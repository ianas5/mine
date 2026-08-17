"""PCCM Phase-5 analytical oracle — pure, plain-data, Linux-side.

Defines the deterministic and analytical semantics that later Stage-A emission and
later VBA must implement, and produces the values a later Windows harness will
assert. There is exactly one definition of the mathematics, and this is it.

--------------------------------------------------------------------------------
WHAT IS NOT HERE, BY CONSTRUCTION
--------------------------------------------------------------------------------
No Excel object model, no openpyxl, no workbook or file I/O, no COM, no VBA, no
randomness, and no Monte Carlo of any kind. Nothing in this module knows what a
cell, a sheet, a ListObject or an address is. A static test proves it.

--------------------------------------------------------------------------------
TWO LAYERS, MIRRORING THE LATER VBA BOUNDARY
--------------------------------------------------------------------------------
    resolution / validation   plain data in, resolved numbers out   -> modCalcResolve
    numerical kernel          resolved numbers in, results out      -> modCalcFactors
                                                                     + modCalcAnalytical

The split is not decoration. The resolution layer is the only place that knows a
currency can be missing or a profile incomplete; the kernel receives numbers that
have already been proven usable, so its failures can only ever be
representability failures. Keeping them apart is what makes the later refusal
classification honest.

`calc_numeric.py` holds the primitives and the failure vocabulary; this module
holds the model, the resolution rules and the analytical mathematics.

--------------------------------------------------------------------------------
IDENTITY IS THE PERMANENT ID, AND SO IS THE EVALUATION ORDER
--------------------------------------------------------------------------------
Profiling weights are carried WITH their driver and keyed by permanent ID. Row
position is not data and is never an input here.

That is not enough on its own. Floating-point addition is not associative, so
accumulating the same contributions in a different sequence can give a different
total - and row order IS excluded from the calculation fingerprint, so two
workbooks with the SAME fingerprint could otherwise disagree on the answer purely
because someone sorted a ListObject.

Every accumulation therefore runs in a CANONICAL ORDER: ascending permanent ID,
ordinal on UTF-16 code units, using the same comparison the fingerprint uses. The
mathematics is unchanged; only the evaluation sequence is pinned.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Mapping, Sequence

from .calc_fingerprint import utf16_sort_key
from .calc_numeric import (
    CalculationRefusal,
    ModelInputRefusal,
    NumericalRangeRefusal,
    OracleError,
    OracleInvariantError,
    allowance_from_scaled,
    beta_pert_mean,
    compound_inflation_factors,
    discount_factor_series,
    is_usable_double,
    midpoint,
    safe_accumulate,
    safe_add,
    safe_multiply,
    safe_product,
    safe_subtract,
    scaled_magnitude,
    triangular_mean,
)

__all__ = [
    "AnalyticalTotals",
    "AnnualRow",
    "AppliedTimeline",
    "CalculationModel",
    "CalculationRefusal",
    "CalculationResult",
    "CostDriver",
    "DistributionKind",
    "DriverFactors",
    "DriverKind",
    "FxRow",
    "IdentityCheck",
    "InflationFactorRow",
    "ModelInputRefusal",
    "NumericalRangeRefusal",
    "OracleError",
    "OracleInvariantError",
    "ReconciliationMagnitudes",
    "RiskDriver",
    "Tolerances",
    "assert_reconciled",
    "calculate",
    "central_basis_label",
    "central_value",
    "expected_value",
    "canonical_order",
    "precomputed_factors",
    "reconcile",
    "resolve_fx",
    "resolve_inflation",
]

REPORTING_CURRENCY = "SAR"
"""The reporting currency, a locked model invariant of `input_contract.yaml`. Its
rate is required to be exactly 1 in every model, referenced or not."""


class DistributionKind(Enum):
    """The kernel's INTERNAL shape vocabulary.

    Deliberately not the accepted display names: this enum says what the
    mathematics does, and nothing here is a second authority over which
    distributions the model offers.
    """

    THREE_POINT_TRIANGULAR = "three-point triangular"
    THREE_POINT_BETA_PERT = "three-point Beta-PERT, lambda 4"
    TWO_POINT_UNIFORM = "two-point uniform"


_DISTRIBUTION_ADAPTER: dict[str, DistributionKind] = {
    "Triangular": DistributionKind.THREE_POINT_TRIANGULAR,
    "Beta-PERT": DistributionKind.THREE_POINT_BETA_PERT,
    "Uniform": DistributionKind.TWO_POINT_UNIFORM,
}
"""ADAPTER, NOT AUTHORITY.

The master list of distributions is owned by `input_contract.yaml`
(`config_tables.distributions`) and is not restated in `calc_contract.yaml` or
here. This mapping only says which internal shape each accepted name selects; a
name absent from it is refused as an invalid distribution, exactly as an unknown
name should be. If the upstream list ever changes, this adapter fails loudly
rather than silently accepting or silently dropping a distribution.
"""

CENTRAL_BASIS_ML = "ML"
CENTRAL_BASIS_MIDPOINT = "Midpoint"


class DriverKind(Enum):
    COST_LINE = "Cost Line"
    RISK = "Risk"


# ---------------------------------------------------------------------------
# The pure data model
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class AppliedTimeline:
    """The APPLIED triple. Entered-but-not-applied values never reach the oracle."""

    base_year: int
    start_year: int
    duration: int

    @property
    def last_year(self) -> int:
        return self.start_year + self.duration - 1

    def project_years(self) -> tuple[tuple[int, int], ...]:
        """`(project index from 1, calendar year)` for each applied project year."""
        return tuple(
            (index, self.start_year + index - 1) for index in range(1, self.duration + 1)
        )


@dataclass(frozen=True)
class FxRow:
    """One `tblFXRates` row, as entered.

    `rate` is deliberately loosely typed: a row may legitimately be blank
    (`None`) or hold something non-numeric, and the resolution layer's job is to
    refuse those - but only if the currency is referenced.
    """

    currency: str
    rate: object


@dataclass(frozen=True)
class _DriverBase:
    permanent_id: str
    distribution: str
    currency: str
    inflation_profile: str
    min_value: object
    most_likely: object
    max_value: object
    profile_weights: tuple[object, ...]
    """One weight per APPLIED project year, in project-index order, keyed to this
    driver by identity rather than by row. `None` is a blank cell and is refused;
    numeric zero is a legitimate weight."""


@dataclass(frozen=True)
class CostDriver(_DriverBase):
    quantity: object = 1.0


@dataclass(frozen=True)
class RiskDriver(_DriverBase):
    probability: object = 1.0


@dataclass(frozen=True)
class CalculationModel:
    """Everything the oracle needs, as plain data.

    `fx_rows` is a SEQUENCE, not a mapping, so a duplicated currency can be
    represented and refused. `inflation_rates` maps profile name to
    calendar year to rate, where a missing year and a `None` rate are different
    facts and both are refused for a referenced profile.
    """

    timeline: AppliedTimeline
    discount_rate: object
    fx_rows: tuple[FxRow, ...] = ()
    inflation_rates: Mapping[str, Mapping[int, object]] = field(default_factory=dict)
    cost_drivers: tuple[CostDriver, ...] = ()
    risk_drivers: tuple[RiskDriver, ...] = ()


@dataclass(frozen=True)
class Tolerances:
    """Numerical tolerances, PASSED IN.

    The values are owned by `spec/calc_contract.yaml`. The pure numerical layer
    never reads YAML or any other file; a caller at the adapter boundary supplies
    the validated values.
    """

    profiling_sum_absolute: float
    identity_absolute_floor: float
    identity_relative_coefficient: float
    conditioning_scale_floor: float = 1.0


@dataclass(frozen=True)
class InflationFactorRow:
    """One audit row. `annual_rate` is `None` for the base year - a
    model-controlled blank, because no rate exists for it to carry."""

    profile: str
    calendar_year: int
    annual_rate: float | None
    cumulative_factor: float


@dataclass(frozen=True)
class DriverFactors:
    """The per-driver audit record. Fields that do not apply to a kind are `None`,
    never zero and never reused - the same rule the `_Calc` schema enforces."""

    permanent_id: str
    driver_kind: DriverKind
    distribution: str
    central_basis: str
    currency: str
    fx_to_sar: float
    inflation_profile: str
    quantity: float | None
    probability: float | None
    central_value: float | None
    mean_value: float
    knom: float
    kpv: float
    deterministic_nominal: float | None
    deterministic_pv: float | None
    mean_basis_nominal: float | None
    mean_basis_pv: float | None
    uncertainty_mean_shift_nominal: float | None
    uncertainty_mean_shift_pv: float | None
    expected_risk_nominal: float | None
    expected_risk_pv: float | None
    weights: tuple[float, ...]


@dataclass(frozen=True)
class AnnualRow:
    project_index: int
    calendar_year: int
    base_cost_nominal: float
    expected_risk_nominal: float
    total_nominal: float
    base_cost_pv: float
    expected_risk_pv: float
    total_pv: float


@dataclass(frozen=True)
class AnalyticalTotals:
    a_nom: float
    a_pv: float
    b_nom: float
    b_pv: float
    c_nom: float
    c_pv: float
    d_nom: float
    d_pv: float
    e_nom: float
    e_pv: float


@dataclass(frozen=True)
class ReconciliationMagnitudes:
    """Conditioning magnitudes, captured WHILE the contributions are accumulated.

    Every field is already multiplied by the relative coefficient, so the raw sum
    of contributions - which can exceed Double where the tolerance does not - is
    never formed.

    WHY UNDERLYING CONTRIBUTIONS AND NOT AGGREGATES. Plan §15 requires the scale
    to reflect "the magnitude of the arithmetic performed, not the magnitude of
    its net result". Summing `|A| + |B| + |C|` measures the *results*, and so does
    summing `|annual aggregate_y|`: both are already-cancelled numbers. A model can
    run its accumulation through `1e17` and end at `A = 32`, or process `2e16` of
    annual contributions and leave an annual aggregate of `1`. Conditioning on
    either would size the tolerance by what survived rather than by what happened,
    and report ordinary Double rounding as a bookkeeping mismatch.

    These are therefore sums over the PER-DRIVER and PER-DRIVER-PER-YEAR
    contributions, before any aggregation. This is internal reconciliation
    metadata; it is NOT a `_Calc` audit column and is not written anywhere.
    """

    relative_coefficient: float
    a_nom: float = 0.0
    a_pv: float = 0.0
    b_nom: float = 0.0
    b_pv: float = 0.0
    c_nom: float = 0.0
    c_pv: float = 0.0
    d_nom: float = 0.0
    d_pv: float = 0.0
    e_nom: float = 0.0
    e_pv: float = 0.0
    annual_base_nom: float = 0.0
    annual_base_pv: float = 0.0
    annual_risk_nom: float = 0.0
    annual_risk_pv: float = 0.0
    annual_total_nom: float = 0.0
    annual_total_pv: float = 0.0


@dataclass(frozen=True)
class CalculationResult:
    totals: AnalyticalTotals
    annual: tuple[AnnualRow, ...]
    drivers: tuple[DriverFactors, ...]
    inflation_factors: tuple[InflationFactorRow, ...]
    discount_factors: Mapping[int, float]
    resolved_fx: Mapping[str, float]
    magnitudes: ReconciliationMagnitudes


@dataclass(frozen=True)
class IdentityCheck:
    name: str
    left: float
    right: float
    difference: float
    allowance: float

    @property
    def holds(self) -> bool:
        return abs(self.difference) <= self.allowance


# ---------------------------------------------------------------------------
# Layer 1 - resolution and validation
# ---------------------------------------------------------------------------
def canonical_order(drivers: Sequence[_DriverBase]) -> tuple[_DriverBase, ...]:
    """Ascending permanent ID, ordinal on UTF-16 code units.

    THE CANONICAL COMPUTATIONAL ORDER. Floating-point addition is not associative,
    so the sequence in which contributions are accumulated is part of the answer,
    not an implementation detail. Row order is deliberately excluded from the
    calculation fingerprint, so two workbooks with the same fingerprint must not be
    able to produce different totals because someone sorted a ListObject.

    The comparison is imported from `calc_fingerprint` rather than reimplemented:
    the fingerprint already had to define ordinal UTF-16 ordering to match VBA's
    `StrComp(..., vbBinaryCompare)`, and two copies of an ordering rule are two
    chances to disagree.

    This is NOT magnitude ordering. Magnitude-balanced evaluation exists only as
    `safe_product`'s tier-2 rescue for a single short product, and never decides
    the order drivers are processed in.
    """
    return tuple(sorted(drivers, key=lambda driver: utf16_sort_key(driver.permanent_id)))


def _canonical_names(names: Sequence[str]) -> tuple[str, ...]:
    """Deterministic ordering for discovered currencies and profiles.

    Their order is observable - it decides the order of the inflation-factor audit
    rows - so it is pinned by the same rule rather than left to whichever driver
    happened to be first in the worksheet.
    """
    return tuple(sorted(set(names), key=utf16_sort_key))


def _numeric(value: object, where: str) -> float:
    """Accept a real number; refuse blank, text and booleans alike.

    `bool` is excluded on purpose: `True` is arithmetically 1 in Python and would
    silently become a quantity or a rate.
    """
    if value is None:
        raise ModelInputRefusal(f"{where}: value is blank")
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ModelInputRefusal(f"{where}: value {value!r} is not numeric")
    if not is_usable_double(value):
        raise ModelInputRefusal(f"{where}: value {value!r} is not a usable Double")
    return float(value)


def _referenced_currencies(model: CalculationModel) -> tuple[str, ...]:
    """Built from the IDENTIFIED DRIVERS, before `tblFXRates` is touched.

    That ordering of operations is the whole referenced-only rule: a Config row
    for a currency nobody uses cannot block a valid model, because resolution
    never asks about it. The RESULT is returned in canonical order, because it is
    observable downstream.
    """
    return _canonical_names(
        [driver.currency for driver in (*model.cost_drivers, *model.risk_drivers)]
    )


def _referenced_profiles(model: CalculationModel) -> tuple[str, ...]:
    return _canonical_names(
        [driver.inflation_profile for driver in (*model.cost_drivers, *model.risk_drivers)]
    )


def resolve_fx(model: CalculationModel) -> dict[str, float]:
    """Resolve REFERENCED currencies only, plus the global SAR invariant.

    SAR is checked in every model, including one that references no currency at
    all and one with no drivers: the reporting currency identity is a global
    invariant, not a per-driver question.
    """
    rows_by_currency: dict[str, list[FxRow]] = {}
    for row in model.fx_rows:
        rows_by_currency.setdefault(row.currency, []).append(row)

    # --- global invariant: SAR = 1 -----------------------------------------
    sar_rows = rows_by_currency.get(REPORTING_CURRENCY, [])
    if len(sar_rows) != 1:
        raise ModelInputRefusal(
            f"FX: the reporting currency {REPORTING_CURRENCY!r} must appear exactly once in "
            f"the FX table, found {len(sar_rows)}. This is a global invariant and applies "
            "whether or not any driver references it."
        )
    sar_rate = _numeric(sar_rows[0].rate, f"FX rate for {REPORTING_CURRENCY!r}")
    if sar_rate != 1.0:
        raise ModelInputRefusal(
            f"FX: the reporting currency {REPORTING_CURRENCY!r} must resolve to exactly 1, "
            f"found {sar_rate!r}"
        )

    resolved = {REPORTING_CURRENCY: 1.0}
    for currency in _referenced_currencies(model):
        if currency == REPORTING_CURRENCY:
            continue
        rows = rows_by_currency.get(currency, [])
        if not rows:
            raise ModelInputRefusal(
                f"FX: currency {currency!r} is referenced by at least one driver but has no "
                "row in the FX table"
            )
        if len(rows) > 1:
            raise ModelInputRefusal(
                f"FX: currency {currency!r} is referenced by at least one driver and appears "
                f"{len(rows)} times in the FX table; exactly one rate must resolve"
            )
        rate = _numeric(rows[0].rate, f"FX rate for referenced currency {currency!r}")
        if rate <= 0.0:
            raise ModelInputRefusal(
                f"FX: referenced currency {currency!r} has rate {rate!r}; an FX rate must be "
                "strictly positive"
            )
        resolved[currency] = rate
    return resolved


def resolve_inflation(model: CalculationModel) -> dict[str, dict[int, float]]:
    """Resolve REFERENCED profiles only, over `BaseYear+1 .. LastProjectYear`.

    A missing required year is refused and never treated as zero: the inflation
    grid is seeded blank precisely so an unmade assumption cannot be fabricated
    as 0%.
    """
    timeline = model.timeline
    required_years = range(timeline.base_year + 1, timeline.last_year + 1)
    resolved: dict[str, dict[int, float]] = {}

    for profile in _referenced_profiles(model):
        if profile not in model.inflation_rates:
            raise ModelInputRefusal(
                f"inflation: profile {profile!r} is referenced by at least one driver but is "
                "not present in the inflation table"
            )
        entered = model.inflation_rates[profile]
        rates: dict[int, float] = {}
        for year in required_years:
            where = f"inflation profile {profile!r}, calendar year {year}"
            if year not in entered:
                raise ModelInputRefusal(f"{where}: required rate is missing")
            rate = _numeric(entered[year], where)
            if 1.0 + rate <= 0.0:
                raise ModelInputRefusal(
                    f"{where}: rate {rate!r} gives 1 + rate <= 0; a rate of -100% or lower "
                    "is refused"
                )
            rates[year] = rate
        resolved[profile] = rates
    return resolved


def _resolve_weights(
    driver: _DriverBase, timeline: AppliedTimeline, tolerances: Tolerances
) -> tuple[float, ...]:
    """Profiling weights for ONE driver, by permanent ID.

    Blank is refused and is not zero (D4). Numeric zero is a legitimate weight: a
    driver may genuinely spend nothing in a given year, and refusing that would
    invent a business rule no contract states.
    """
    where = f"profiling for driver {driver.permanent_id!r}"
    if len(driver.profile_weights) != timeline.duration:
        raise ModelInputRefusal(
            f"{where}: {len(driver.profile_weights)} weights declared for an applied duration "
            f"of {timeline.duration} project years"
        )
    weights: list[float] = []
    for offset, raw in enumerate(driver.profile_weights):
        index = offset + 1
        cell = f"{where}, project year {index}"
        if raw is None:
            raise ModelInputRefusal(
                f"{cell}: profiling cell is blank. A blank is not zero - it is an unmade "
                "assumption, and it is refused rather than fabricated."
            )
        weights.append(_numeric(raw, cell))

    total = 0.0
    for offset, weight in enumerate(weights):
        total = safe_accumulate(total, weight, f"{where}, project year {offset + 1}")
    if abs(total - 1.0) > tolerances.profiling_sum_absolute:
        raise ModelInputRefusal(
            f"{where}: weights sum to {total!r}, which is not 100% within "
            f"{tolerances.profiling_sum_absolute!r}"
        )
    return tuple(weights)


def _resolve_distribution(driver: _DriverBase) -> DistributionKind:
    kind = _DISTRIBUTION_ADAPTER.get(driver.distribution)
    if kind is None:
        raise ModelInputRefusal(
            f"driver {driver.permanent_id!r}: distribution {driver.distribution!r} is not one "
            f"of the accepted distributions {sorted(_DISTRIBUTION_ADAPTER)}"
        )
    return kind


def _resolve_three_point(
    driver: _DriverBase, kind: DistributionKind
) -> tuple[float, float | None, float]:
    """Min / ML / Max with the locked ordering rules, and nothing else.

    NO POSITIVITY RULE IS INVENTED. No locked contract requires Min, ML or Max to
    be non-negative, so an ordered set of negative values is a valid model.
    """
    where = f"driver {driver.permanent_id!r}"
    minimum = _numeric(driver.min_value, f"{where}: Min")
    maximum = _numeric(driver.max_value, f"{where}: Max")

    if kind is DistributionKind.TWO_POINT_UNIFORM:
        # D1: a populated ML is ACCEPTED and IGNORED. Uniform is a two-point
        # distribution; the cell may hold a leftover value from another choice of
        # distribution, and refusing it would block a valid model.
        if minimum > maximum:
            raise ModelInputRefusal(
                f"{where}: Uniform requires Min <= Max, got Min {minimum!r} and Max {maximum!r}"
            )
        return minimum, None, maximum

    most_likely = _numeric(driver.most_likely, f"{where}: Most Likely")
    if not (minimum <= most_likely <= maximum):
        raise ModelInputRefusal(
            f"{where}: {driver.distribution} requires Min <= Most Likely <= Max, got "
            f"{minimum!r}, {most_likely!r}, {maximum!r}"
        )
    return minimum, most_likely, maximum


# ---------------------------------------------------------------------------
# Layer 2 - the numerical kernel
# ---------------------------------------------------------------------------
def central_value(kind: DistributionKind, minimum: float, most_likely: float | None,
                  maximum: float) -> float:
    """The DETERMINISTIC central value — risks excluded, never called "mean"."""
    if kind is DistributionKind.TWO_POINT_UNIFORM:
        return midpoint(minimum, maximum)
    assert most_likely is not None  # guaranteed by _resolve_three_point
    return most_likely


def expected_value(kind: DistributionKind, minimum: float, most_likely: float | None,
                   maximum: float) -> float:
    """The distribution mean, always by the stable form (§19.2)."""
    if kind is DistributionKind.TWO_POINT_UNIFORM:
        return midpoint(minimum, maximum)
    assert most_likely is not None
    if kind is DistributionKind.THREE_POINT_TRIANGULAR:
        return triangular_mean(minimum, most_likely, maximum)
    return beta_pert_mean(minimum, most_likely, maximum)


def central_basis_label(kind: DistributionKind) -> str:
    """An explicit audit field: an auditor must not have to infer the basis."""
    return (
        CENTRAL_BASIS_MIDPOINT
        if kind is DistributionKind.TWO_POINT_UNIFORM
        else CENTRAL_BASIS_ML
    )


def precomputed_factors(
    fx_rate: float,
    weights: Sequence[float],
    inflation_by_index: Sequence[float],
    discount_by_index: Sequence[float],
    where: str,
) -> tuple[float, float]:
    """`Knom` and `Kpv` for one driver.

    ```
    Knom = FX * SUM_y ( w_y * infl_y )
    Kpv  = FX * SUM_y ( w_y * infl_y * disc_y )
    ```

    QUANTITY AND PROBABILITY ARE DELIBERATELY ABSENT. Probability is replaced by a
    Bernoulli draw in Monte Carlo and must not be folded in here; Quantity is a
    per-driver multiplier, not a factor of the escalation path. Folding either in
    would double-count them at the contribution step.
    """
    nominal = 0.0
    present = 0.0
    for offset, weight in enumerate(weights):
        index = offset + 1
        year_where = f"{where}, project year {index}"
        nominal = safe_accumulate(
            nominal, safe_product([weight, inflation_by_index[offset]], year_where), year_where
        )
        present = safe_accumulate(
            present,
            safe_product(
                [weight, inflation_by_index[offset], discount_by_index[offset]], year_where
            ),
            year_where,
        )
    return (
        safe_multiply(fx_rate, nominal, f"{where}: Knom"),
        safe_multiply(fx_rate, present, f"{where}: Kpv"),
    )


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def calculate(model: CalculationModel, tolerances: Tolerances) -> CalculationResult:
    """Resolve, validate, then calculate — entirely in memory.

    Nothing is written anywhere; this function has no side effects at all. The
    later transactional write-back is a separate concern and is not modelled here.
    """
    timeline = model.timeline
    if timeline.duration < 1:
        raise ModelInputRefusal(
            f"timeline: applied duration is {timeline.duration}; at least one project year "
            "is required"
        )
    if timeline.base_year > timeline.start_year:
        raise ModelInputRefusal(
            f"timeline: Base Year {timeline.base_year} is after Start Year "
            f"{timeline.start_year}; the price base cannot postdate the project"
        )

    discount_rate = _numeric(model.discount_rate, "discount rate")
    if 1.0 + discount_rate <= 0.0:
        raise ModelInputRefusal(
            f"discount rate: {discount_rate!r} gives 1 + r <= 0; a rate of -100% or lower "
            "is refused"
        )

    resolved_fx = resolve_fx(model)
    resolved_rates = resolve_inflation(model)
    discounts = discount_factor_series(discount_rate, timeline.duration)

    # --- inflation factors, per referenced profile, over the audit span -----
    factors_by_profile: dict[str, dict[int, float]] = {}
    inflation_rows: list[InflationFactorRow] = []
    for profile in _referenced_profiles(model):
        factors = compound_inflation_factors(
            timeline.base_year, timeline.last_year, resolved_rates[profile], profile
        )
        factors_by_profile[profile] = factors
        for year in range(timeline.base_year, timeline.last_year + 1):
            inflation_rows.append(
                InflationFactorRow(
                    profile=profile,
                    calendar_year=year,
                    # The Base-Year row carries a BLANK rate: it is a
                    # model-controlled audit output, and no rate exists for it.
                    annual_rate=None if year == timeline.base_year
                    else resolved_rates[profile][year],
                    cumulative_factor=factors[year],
                )
            )

    project_years = timeline.project_years()
    discount_by_index = [discounts[index] for index, _ in project_years]

    drivers: list[DriverFactors] = []

    # CANONICAL ORDER, not worksheet order: everything below accumulates, and
    # floating-point addition is not associative.
    ordered_costs = canonical_order(model.cost_drivers)
    ordered_risks = canonical_order(model.risk_drivers)

    # --- cost lines ---------------------------------------------------------
    for driver in ordered_costs:
        where = f"cost line {driver.permanent_id!r}"
        kind = _resolve_distribution(driver)
        minimum, most_likely, maximum = _resolve_three_point(driver, kind)
        weights = _resolve_weights(driver, timeline, tolerances)

        quantity = _numeric(driver.quantity, f"{where}: Quantity")
        if quantity <= 0.0:
            raise ModelInputRefusal(
                f"{where}: Quantity is {quantity!r}; Quantity must be numeric and strictly "
                "positive for a calculation to be meaningful"
            )

        fx_rate = resolved_fx[driver.currency]
        inflation = [
            factors_by_profile[driver.inflation_profile][year] for _, year in project_years
        ]
        knom, kpv = precomputed_factors(fx_rate, weights, inflation, discount_by_index, where)

        central = central_value(kind, minimum, most_likely, maximum)
        mean = expected_value(kind, minimum, most_likely, maximum)
        shift = safe_subtract(mean, central, f"{where}: uncertainty mean shift")

        drivers.append(
            DriverFactors(
                permanent_id=driver.permanent_id,
                driver_kind=DriverKind.COST_LINE,
                distribution=driver.distribution,
                central_basis=central_basis_label(kind),
                currency=driver.currency,
                fx_to_sar=fx_rate,
                inflation_profile=driver.inflation_profile,
                quantity=quantity,
                probability=None,
                central_value=central,
                mean_value=mean,
                knom=knom,
                kpv=kpv,
                deterministic_nominal=safe_product([central, quantity, knom], f"{where}: A nom"),
                deterministic_pv=safe_product([central, quantity, kpv], f"{where}: A pv"),
                mean_basis_nominal=safe_product([mean, quantity, knom], f"{where}: C nom"),
                mean_basis_pv=safe_product([mean, quantity, kpv], f"{where}: C pv"),
                uncertainty_mean_shift_nominal=safe_product(
                    [shift, quantity, knom], f"{where}: B nom"
                ),
                uncertainty_mean_shift_pv=safe_product([shift, quantity, kpv], f"{where}: B pv"),
                expected_risk_nominal=None,
                expected_risk_pv=None,
                weights=weights,
            )
        )

    # --- risks --------------------------------------------------------------
    for driver in ordered_risks:
        where = f"risk {driver.permanent_id!r}"
        kind = _resolve_distribution(driver)
        minimum, most_likely, maximum = _resolve_three_point(driver, kind)
        weights = _resolve_weights(driver, timeline, tolerances)

        probability = _numeric(driver.probability, f"{where}: Probability")
        if not 0.0 <= probability <= 1.0:
            raise ModelInputRefusal(
                f"{where}: Probability is {probability!r}; it must be a fraction in [0, 1]"
            )

        fx_rate = resolved_fx[driver.currency]
        inflation = [
            factors_by_profile[driver.inflation_profile][year] for _, year in project_years
        ]
        knom, kpv = precomputed_factors(fx_rate, weights, inflation, discount_by_index, where)

        severity = expected_value(kind, minimum, most_likely, maximum)

        drivers.append(
            DriverFactors(
                permanent_id=driver.permanent_id,
                driver_kind=DriverKind.RISK,
                distribution=driver.distribution,
                central_basis=central_basis_label(kind),
                currency=driver.currency,
                fx_to_sar=fx_rate,
                inflation_profile=driver.inflation_profile,
                quantity=None,
                probability=probability,
                central_value=None,
                mean_value=severity,
                knom=knom,
                kpv=kpv,
                deterministic_nominal=None,
                deterministic_pv=None,
                mean_basis_nominal=None,
                mean_basis_pv=None,
                uncertainty_mean_shift_nominal=None,
                uncertainty_mean_shift_pv=None,
                expected_risk_nominal=safe_product(
                    [probability, severity, knom], f"{where}: D nom"
                ),
                expected_risk_pv=safe_product([probability, severity, kpv], f"{where}: D pv"),
                weights=weights,
            )
        )

    totals, totals_magnitudes = _accumulate_totals(drivers, tolerances)
    annual, annual_magnitudes = _annual_series(
        ordered_costs, ordered_risks, drivers, factors_by_profile, discounts,
        project_years, tolerances,
    )
    return CalculationResult(
        totals=totals,
        annual=annual,
        drivers=tuple(drivers),
        inflation_factors=tuple(inflation_rows),
        discount_factors=dict(discounts),
        resolved_fx=dict(resolved_fx),
        magnitudes=_merge_magnitudes(totals_magnitudes, annual_magnitudes),
    )


def _merge_magnitudes(
    totals: ReconciliationMagnitudes, annual: ReconciliationMagnitudes
) -> ReconciliationMagnitudes:
    """Combine the two capture passes into one record."""
    if totals.relative_coefficient != annual.relative_coefficient:
        raise OracleInvariantError("conditioning magnitudes captured at different coefficients")
    return ReconciliationMagnitudes(
        relative_coefficient=totals.relative_coefficient,
        a_nom=totals.a_nom, a_pv=totals.a_pv,
        b_nom=totals.b_nom, b_pv=totals.b_pv,
        c_nom=totals.c_nom, c_pv=totals.c_pv,
        d_nom=totals.d_nom, d_pv=totals.d_pv,
        e_nom=totals.e_nom, e_pv=totals.e_pv,
        annual_base_nom=annual.annual_base_nom, annual_base_pv=annual.annual_base_pv,
        annual_risk_nom=annual.annual_risk_nom, annual_risk_pv=annual.annual_risk_pv,
        annual_total_nom=annual.annual_total_nom, annual_total_pv=annual.annual_total_pv,
    )


def _accumulate_totals(
    drivers: Sequence[DriverFactors], tolerances: Tolerances
) -> tuple[AnalyticalTotals, ReconciliationMagnitudes]:
    """Five measures, each accumulated in its OWN pass, in canonical driver order.

    `B = C - A` and `E = C + D` are NOT the calculation path. They are
    reconciliation identities, and an identity computed by definition checks
    nothing. Accumulating each measure independently is what makes I1 and I2 real.

    Each accumulation is checked at every driver, so a failure names the driver
    rather than reporting that a total came out infinite. Alongside each
    contribution its scaled absolute magnitude is captured, so the conditioning
    scale measures the arithmetic THAT WAS PERFORMED rather than the total that
    survived it.
    """
    a_nom = a_pv = b_nom = b_pv = c_nom = c_pv = d_nom = d_pv = e_nom = e_pv = 0.0
    mag: dict[str, float] = {
        key: 0.0 for key in
        ("a_nom", "a_pv", "b_nom", "b_pv", "c_nom", "c_pv", "d_nom", "d_pv", "e_nom", "e_pv")
    }
    coefficient = tolerances.identity_relative_coefficient

    def record(key: str, value: float, where: str) -> None:
        mag[key] = scaled_magnitude(mag[key], value, coefficient, where)

    for driver in drivers:
        tag = f"totals, driver {driver.permanent_id!r}"
        if driver.driver_kind is DriverKind.COST_LINE:
            a_nom = safe_accumulate(a_nom, driver.deterministic_nominal, f"{tag}: A nom")
            a_pv = safe_accumulate(a_pv, driver.deterministic_pv, f"{tag}: A pv")
            b_nom = safe_accumulate(
                b_nom, driver.uncertainty_mean_shift_nominal, f"{tag}: B nom"
            )
            b_pv = safe_accumulate(b_pv, driver.uncertainty_mean_shift_pv, f"{tag}: B pv")
            c_nom = safe_accumulate(c_nom, driver.mean_basis_nominal, f"{tag}: C nom")
            c_pv = safe_accumulate(c_pv, driver.mean_basis_pv, f"{tag}: C pv")
            record("a_nom", driver.deterministic_nominal, f"{tag}: |A| nom")
            record("a_pv", driver.deterministic_pv, f"{tag}: |A| pv")
            record("b_nom", driver.uncertainty_mean_shift_nominal, f"{tag}: |B| nom")
            record("b_pv", driver.uncertainty_mean_shift_pv, f"{tag}: |B| pv")
            record("c_nom", driver.mean_basis_nominal, f"{tag}: |C| nom")
            record("c_pv", driver.mean_basis_pv, f"{tag}: |C| pv")
        else:
            d_nom = safe_accumulate(d_nom, driver.expected_risk_nominal, f"{tag}: D nom")
            d_pv = safe_accumulate(d_pv, driver.expected_risk_pv, f"{tag}: D pv")
            record("d_nom", driver.expected_risk_nominal, f"{tag}: |D| nom")
            record("d_pv", driver.expected_risk_pv, f"{tag}: |D| pv")

    # E is accumulated in its OWN pass over the same contributions, not derived
    # from C and D. Two independent journeys to the same number are what I2 tests.
    for driver in drivers:
        tag = f"totals, driver {driver.permanent_id!r}: E"
        if driver.driver_kind is DriverKind.COST_LINE:
            e_nom = safe_accumulate(e_nom, driver.mean_basis_nominal, f"{tag} nom")
            e_pv = safe_accumulate(e_pv, driver.mean_basis_pv, f"{tag} pv")
            record("e_nom", driver.mean_basis_nominal, f"{tag} |E| nom")
            record("e_pv", driver.mean_basis_pv, f"{tag} |E| pv")
        else:
            e_nom = safe_accumulate(e_nom, driver.expected_risk_nominal, f"{tag} nom")
            e_pv = safe_accumulate(e_pv, driver.expected_risk_pv, f"{tag} pv")
            record("e_nom", driver.expected_risk_nominal, f"{tag} |E| nom")
            record("e_pv", driver.expected_risk_pv, f"{tag} |E| pv")

    totals = AnalyticalTotals(a_nom, a_pv, b_nom, b_pv, c_nom, c_pv, d_nom, d_pv, e_nom, e_pv)
    return totals, ReconciliationMagnitudes(relative_coefficient=coefficient, **mag)


def _annual_series(
    ordered_costs: Sequence[CostDriver],
    ordered_risks: Sequence[RiskDriver],
    drivers: Sequence[DriverFactors],
    factors_by_profile: Mapping[str, Mapping[int, float]],
    discounts: Mapping[int, float],
    project_years: Sequence[tuple[int, int]],
    tolerances: Tolerances,
) -> tuple[tuple[AnnualRow, ...], ReconciliationMagnitudes]:
    """Six series per applied project year, on the MEAN basis, in canonical order.

    Annual Base Cost uses the distribution expected value, not the deterministic
    ML/Midpoint basis: the locked Results requirement is that annual cash flow is
    mean-only, so the deterministic basis has no annual series at all.

    The annual TOTAL is accumulated in its own pass rather than added from the two
    series above it, so I3c and I4c are real checks and not arithmetic identities.

    Conditioning magnitudes are captured PER DRIVER PER YEAR, before the row
    aggregate exists. Conditioning on the aggregate would be conditioning on a
    number that has already cancelled - the annual aggregate can be `1` where the
    contributions that produced it were `1e16` apart.
    """
    by_id = {driver.permanent_id: driver for driver in drivers}
    rows: list[AnnualRow] = []
    coefficient = tolerances.identity_relative_coefficient
    mag: dict[str, float] = {
        key: 0.0 for key in
        ("annual_base_nom", "annual_base_pv", "annual_risk_nom", "annual_risk_pv",
         "annual_total_nom", "annual_total_pv")
    }

    def record(key: str, value: float, where: str) -> None:
        mag[key] = scaled_magnitude(mag[key], value, coefficient, where)

    for offset, (index, calendar_year) in enumerate(project_years):
        base_nom = risk_nom = total_nom = 0.0
        base_pv = risk_pv = total_pv = 0.0
        discount = discounts[index]

        for driver in ordered_costs:
            resolved = by_id[driver.permanent_id]
            where = f"annual year {calendar_year}, cost line {driver.permanent_id!r}"
            infl = factors_by_profile[driver.inflation_profile][calendar_year]
            weight = resolved.weights[offset]
            nominal = safe_product(
                [resolved.mean_value, resolved.quantity, resolved.fx_to_sar, weight, infl], where
            )
            present = safe_product([nominal, discount], f"{where} PV")
            base_nom = safe_accumulate(base_nom, nominal, f"{where}: base nominal")
            base_pv = safe_accumulate(base_pv, present, f"{where}: base PV")
            total_nom = safe_accumulate(total_nom, nominal, f"{where}: total nominal")
            total_pv = safe_accumulate(total_pv, present, f"{where}: total PV")
            record("annual_base_nom", nominal, f"{where}: |base| nominal")
            record("annual_base_pv", present, f"{where}: |base| PV")
            record("annual_total_nom", nominal, f"{where}: |total| nominal")
            record("annual_total_pv", present, f"{where}: |total| PV")

        for driver in ordered_risks:
            resolved = by_id[driver.permanent_id]
            where = f"annual year {calendar_year}, risk {driver.permanent_id!r}"
            infl = factors_by_profile[driver.inflation_profile][calendar_year]
            weight = resolved.weights[offset]
            nominal = safe_product(
                [resolved.probability, resolved.mean_value, resolved.fx_to_sar, weight, infl],
                where,
            )
            present = safe_product([nominal, discount], f"{where} PV")
            risk_nom = safe_accumulate(risk_nom, nominal, f"{where}: risk nominal")
            risk_pv = safe_accumulate(risk_pv, present, f"{where}: risk PV")
            total_nom = safe_accumulate(total_nom, nominal, f"{where}: total nominal")
            total_pv = safe_accumulate(total_pv, present, f"{where}: total PV")
            record("annual_risk_nom", nominal, f"{where}: |risk| nominal")
            record("annual_risk_pv", present, f"{where}: |risk| PV")
            record("annual_total_nom", nominal, f"{where}: |total| nominal")
            record("annual_total_pv", present, f"{where}: |total| PV")

        rows.append(
            AnnualRow(
                project_index=index,
                calendar_year=calendar_year,
                base_cost_nominal=base_nom,
                expected_risk_nominal=risk_nom,
                total_nominal=total_nom,
                base_cost_pv=base_pv,
                expected_risk_pv=risk_pv,
                total_pv=total_pv,
            )
        )
    return tuple(rows), ReconciliationMagnitudes(relative_coefficient=coefficient, **mag)


# ---------------------------------------------------------------------------
# Reconciliation
# ---------------------------------------------------------------------------
def reconcile(result: CalculationResult, tolerances: Tolerances) -> tuple[IdentityCheck, ...]:
    """I1 - I5, conditioned on the UNDERLYING CONTRIBUTIONS.

    Plan §15 erratum C1. Every scale sums the scaled absolute magnitudes of the
    per-driver (and per-driver-per-year) contributions that the identity
    accumulated - never the headline totals, and never the annual row aggregates.
    Both of those are already-cancelled numbers, and conditioning on them sizes
    the tolerance by what survived instead of by what happened.

    The magnitudes were captured during accumulation, so they cannot describe a
    different calculation from the one being checked.

    All arithmetic here goes through the safe primitives. A reconciliation that
    quietly produced `inf` and then compared it would be reporting nonsense.
    """
    totals = result.totals
    magnitudes = result.magnitudes
    if magnitudes.relative_coefficient != tolerances.identity_relative_coefficient:
        raise OracleInvariantError(
            f"conditioning magnitudes were captured at coefficient "
            f"{magnitudes.relative_coefficient!r} but reconciliation was asked for "
            f"{tolerances.identity_relative_coefficient!r}; the scales would describe a "
            "different tolerance from the one requested"
        )

    checks: list[IdentityCheck] = []

    def check(name: str, left: float, right: float, scaled_terms: Sequence[float]) -> None:
        scaled = 0.0
        for term in scaled_terms:
            scaled = safe_accumulate(scaled, term, f"{name}: conditioning scale")
        allowance = allowance_from_scaled(
            scaled,
            tolerances.identity_absolute_floor,
            tolerances.identity_relative_coefficient,
            tolerances.conditioning_scale_floor,
        )
        difference = safe_subtract(left, right, f"{name}: difference")
        checks.append(IdentityCheck(name, left, right, difference, allowance))

    check(
        "I1 nominal: A + B = C",
        safe_add(totals.a_nom, totals.b_nom, "I1 nominal"),
        totals.c_nom,
        (magnitudes.a_nom, magnitudes.b_nom, magnitudes.c_nom),
    )
    check(
        "I1 PV: A + B = C",
        safe_add(totals.a_pv, totals.b_pv, "I1 PV"),
        totals.c_pv,
        (magnitudes.a_pv, magnitudes.b_pv, magnitudes.c_pv),
    )
    check(
        "I2 nominal: C + D = E",
        safe_add(totals.c_nom, totals.d_nom, "I2 nominal"),
        totals.e_nom,
        (magnitudes.c_nom, magnitudes.d_nom, magnitudes.e_nom),
    )
    check(
        "I2 PV: C + D = E",
        safe_add(totals.c_pv, totals.d_pv, "I2 PV"),
        totals.e_pv,
        (magnitudes.c_pv, magnitudes.d_pv, magnitudes.e_pv),
    )

    annual = result.annual
    series = (
        ("I3a nominal base", [r.base_cost_nominal for r in annual], totals.c_nom,
         (magnitudes.annual_base_nom, magnitudes.c_nom)),
        ("I3b nominal risk", [r.expected_risk_nominal for r in annual], totals.d_nom,
         (magnitudes.annual_risk_nom, magnitudes.d_nom)),
        ("I3c nominal total", [r.total_nominal for r in annual], totals.e_nom,
         (magnitudes.annual_total_nom, magnitudes.e_nom)),
        ("I4a PV base", [r.base_cost_pv for r in annual], totals.c_pv,
         (magnitudes.annual_base_pv, magnitudes.c_pv)),
        ("I4b PV risk", [r.expected_risk_pv for r in annual], totals.d_pv,
         (magnitudes.annual_risk_pv, magnitudes.d_pv)),
        ("I4c PV total", [r.total_pv for r in annual], totals.e_pv,
         (magnitudes.annual_total_pv, magnitudes.e_pv)),
    )
    for name, values, headline, scaled_terms in series:
        total = 0.0
        for value in values:
            total = safe_accumulate(total, value, name)
        check(name, total, headline, scaled_terms)

    for driver in result.drivers:
        total = 0.0
        for weight in driver.weights:
            total = safe_accumulate(total, weight, f"I5 {driver.permanent_id}")
        checks.append(
            IdentityCheck(
                f"I5 profile sum: {driver.permanent_id}",
                total,
                1.0,
                safe_subtract(total, 1.0, f"I5 {driver.permanent_id}"),
                tolerances.profiling_sum_absolute,
            )
        )
    return tuple(checks)


def assert_reconciled(result: CalculationResult, tolerances: Tolerances) -> None:
    """Raise if any identity fails.

    `OracleInvariantError`, NOT a refusal: the inputs were accepted and the
    calculation ran, so a mismatch here means two independently accumulated
    quantities that must agree do not. That is a defect in the calculation, and
    reporting it to a user as an invalid model would be wrong.
    """
    broken = [check for check in reconcile(result, tolerances) if not check.holds]
    if broken:
        detail = "; ".join(
            f"{c.name}: {c.left!r} vs {c.right!r}, difference {c.difference!r} exceeds "
            f"allowance {c.allowance!r}"
            for c in broken
        )
        raise OracleInvariantError(f"reconciliation failed: {detail}")
