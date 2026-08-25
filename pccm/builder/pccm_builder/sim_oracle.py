#!/usr/bin/env python3
"""PCCM Phase 6 Step-4 pure Python Monte Carlo oracle.

Given one accepted resolved model, an iteration count and an EFFECTIVE seed,
this module answers exactly three questions:

    what nominal and PV total belongs to each iteration,
    what digest identifies those retained totals,
    and what summary statistics follow from them.

It is deterministic, worksheet-free, COM-free, row-order invariant,
stream-independent, transaction-neutral and free of side effects. There is no
module-level mutable simulation state: every current RNG state lives in a local
run context and dies with the call.

--------------------------------------------------------------------------------
IT DOES NOT CONTAIN A CALCULATION MODEL
--------------------------------------------------------------------------------
FX, inflation, profiling, discounting, Quantity validation, Probability
validation and distribution business validation are the accepted Phase-5
oracle's, and they are performed by calling `calculate(...)` ONCE, before the
loop, through its own public entry point. What this module keeps is the small
per-driver residue the hot loop actually needs - `Quantity`, `Probability`,
`Knom`, `Kpv`, the distribution identity, its Min/Most Likely/Max, a prepared
Beta-PERT shape where applicable, and the component stream each driver draws
from. Nothing else survives preparation, and nothing recomputes a Phase-5
quantity inside an iteration.

Min / Most Likely / Max are resolved by the accepted Phase-5 resolvers
themselves - `_resolve_distribution` and `_resolve_three_point` - rather than by
a second coercion here. They are private names, and importing them is a
deliberate trade: the alternative is a duplicate implementation of ordering,
numeric coercion and the D1 "Uniform ignores Most Likely" rule, which is exactly
the second calculation model Step 4 is forbidden to build. A duplicate would be
free to drift; a call cannot.

--------------------------------------------------------------------------------
WHAT IT IS NOT
--------------------------------------------------------------------------------
No `_SimData`. No Results publication. No simulation status, no `run_id`, no
`next_auto_nonce`, no attempt metadata, no workbook fingerprint publication, no
`PCCM_RunSimulation`, no VBA. Step 4 produces pure Python values; the
transactional and reporting boundary is later and is not modelled here.

The engine also reads nothing at run time - no file, no evidence, no workbook.
Contracts and models are supplied to it as already-parsed objects.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Any, Mapping, Sequence

from .calc_fingerprint import (
    Field,
    encode_section,
    fingerprint,
    integer_field,
    number_field,
    text_field,
)
from .calc_numeric import (
    CalculationRefusal,
    safe_product,
    safe_signed_sum,
    safe_subtract,
)
from .calc_oracle import (
    CalculationModel,
    CalculationResult,
    DriverKind,
    Tolerances,
    _DriverBase,
    _resolve_distribution,
    _resolve_three_point,
    calculate,
    canonical_order,
)
from .contract_loader import ContractError, InputContract
from .sim_loader import SimContract
from .sim_rng import (
    COST_KIND,
    RISK_KIND,
    ROLE_OCCURRENCE,
    ROLE_SEVERITY,
    ROLE_VALUE,
    Component,
    RngReference,
    RngState,
)
from .sim_sample import (
    ACCEPTED_FAMILIES,
    FAMILY_BETA_PERT,
    FAMILY_TRIANGULAR,
    FAMILY_UNIFORM,
    PreparedBetaPert,
    bernoulli_occurs,
    prepare_beta_pert,
    sample_prepared_beta,
    sample_triangular,
    sample_uniform,
)
from .sim_stats import MeasureStatistics, SimStatsError, describe

ITERATIONS_INPUT_KEY = "monte_carlo_iterations"
"""The `input_contract.yaml` input whose validation owns the business minimum."""

CONFIDENCE_LEVELS_TABLE_KEY = "confidence_levels"
"""The `input_contract.yaml` config table that owns the selectable ladder."""

_PERCENTILE_LABEL = re.compile(r"^P(\d{1,3})$")


class SimOracleError(ContractError):
    """A simulation the oracle refuses to run, or a contract it cannot read.

    Numerical refusals from the underlying calculation are NOT converted into
    this: `NumericalRangeRefusal` and `ModelInputRefusal` propagate as
    themselves, with the iteration index added to their message, so the accepted
    Phase-5 hierarchy still tells a caller what kind of failure occurred.
    """


# ===========================================================================
# RNG reference binding
# ===========================================================================
def _require_exact_reference(reference: object, where: str) -> RngReference:
    """`RngReference` is the accepted Step-2 oracle, not a plugin interface.

    `isinstance` is deliberately NOT used. An alternate implementation cannot be
    allowed to claim the contract's `rng_version` by inheriting from the
    accepted class and copying its constants - the version identifies a
    generator, and a generator is its behaviour, not its field values.
    """
    if type(reference) is not RngReference:
        raise SimOracleError(
            f"{where}: the RNG reference must be exactly {RngReference.__name__}, got "
            f"{type(reference).__name__}. `RngReference` is the accepted Step-2 oracle "
            "implementation, not an extension point: a subclass or alternate "
            "implementation cannot claim the accepted RNG_VERSION by copying the same "
            "constants, because the constants do not describe its behaviour."
        )
    return reference


def rng_reference_signature(reference: RngReference) -> tuple:
    """A canonical immutable snapshot of every operational field of a reference.

    WHY THIS EXISTS. A prepared model records `rng_version = 1`, and that claim
    has to mean something. Without a binding, a run prepared against the accepted
    contracts could be executed under a reference whose `a12` is one larger,
    producing a completely different digest while still reporting
    `RNG_VERSION = 1`. Two different generators cannot both be version 1.

    IT IS A VALUE, NOT A POINTER. `RngReference` is a frozen dataclass, but
    `role_order` inside it is a live `dict`; keeping the object and calling that
    a binding would let the ordering be rewritten after preparation. Every field
    is therefore copied into tuples here.

    IT IS NOT A HASH AUTHORITY. No digest is computed and no new hash is
    introduced - the comparison is tuple equality, which is exact and needs no
    collision argument.

    Covered: the recurrence constants and normalisation, both jump matrices, the
    AUTO seed cycle, the FIXED seed domain, and the component kind and role
    orders - everything able to change a number or a stream identity.

    WHAT A FIELD SNAPSHOT CANNOT COVER IS BEHAVIOUR, which is why the EXACT type
    is required here too and not merely `isinstance`. A subclass can copy every
    accepted constant and override `next_uniform` to return `0.5` forever; its
    signature would be identical to the accepted reference's while its output is
    a different generator's. Refusing the subclass at the boundary is the only
    check that closes that, because no snapshot of data can see a method body.
    """
    _require_exact_reference(reference, "a reference signature")
    return (
        reference.m1,
        reference.m2,
        reference.a12,
        reference.a13n,
        reference.a21,
        reference.a23n,
        reference.norm,
        tuple(tuple(int(word) for word in row) for row in reference.jump_a1),
        tuple(tuple(int(word) for word in row) for row in reference.jump_a2),
        reference.auto_modulus,
        reference.auto_multiplier,
        reference.nonce_exhausted,
        reference.seed_min,
        reference.seed_max,
        tuple(reference.kind_order),
        tuple(sorted((kind, tuple(roles)) for kind, roles in reference.role_order.items())),
    )


def _binding_mismatch(expected: tuple, actual: tuple) -> str:
    """Which operational fields differ, so a refusal can say what moved."""
    labels = (
        "m1", "m2", "a12", "a13n", "a21", "a23n", "norm",
        "jump.a1_p127", "jump.a2_p127",
        "auto.modulus", "auto.multiplier", "auto.exhausted_value",
        "seed_min", "seed_max", "components.kind_order", "components.role_order",
    )
    return ", ".join(
        label for label, want, got in zip(labels, expected, actual) if want != got
    ) or "an unlabelled field"


# ===========================================================================
# the prepared model
# ===========================================================================
@dataclass(frozen=True)
class PreparedSimulationDriver:
    """One driver, reduced to what a single iteration needs and nothing more.

    No worksheet object, no `ListObject`, no `Range`, no cell address, no
    workbook handle - and no FX rate, weight vector or inflation series either,
    because those have already collapsed into `knom` and `kpv`.

    `most_likely` is `None` for a Uniform. The accepted D1 rule ignores a
    Uniform's Most Likely, so carrying the entered value forward would let a
    meaningless input travel through preparation looking authoritative.
    """

    permanent_id: str
    driver_kind: str
    distribution: str
    minimum: float
    most_likely: float | None
    maximum: float
    quantity: float | None
    probability: float | None
    knom: float
    kpv: float
    beta_shape: PreparedBetaPert | None
    value_stream_index: int
    value_initial_state: RngState
    occurrence_stream_index: int | None
    occurrence_initial_state: RngState | None
    nominal_where: str
    pv_where: str

    @property
    def is_cost_line(self) -> bool:
        return self.driver_kind == COST_KIND


@dataclass(frozen=True)
class PercentileLadder:
    """The reported percentiles, resolved from their owners and not restated.

    `selectable` comes from `input_contract.yaml`; `fixed` from
    `sim_contract.yaml`. `points` is the union in ascending order and is what the
    statistics layer is asked for. Nothing in this module hard-codes a ladder.
    """

    fixed: tuple[str, ...]
    selectable: tuple[str, ...]
    ordered: tuple[str, ...]
    headline: tuple[str, ...]
    points: tuple[tuple[str, float], ...]


@dataclass(frozen=True)
class DeterministicBase:
    """The Phase-5 deterministic base estimate A, for both measures."""

    nominal: float
    pv: float


@dataclass(frozen=True)
class AnalyticalExpectation:
    """The Phase-5 analytical expected total E, for both measures.

    Reporting and cross-check reference only. It is NOT a contingency baseline
    and never enters an iteration.
    """

    nominal: float
    pv: float


@dataclass(frozen=True)
class PreparedSimulationModel:
    """Everything resolved, in memory, before the first random draw."""

    cost_drivers: tuple[PreparedSimulationDriver, ...]
    risk_drivers: tuple[PreparedSimulationDriver, ...]
    iterations: int
    effective_seed: int
    base_state: RngState
    rng_version: int
    sim_method_version: int
    ladder: PercentileLadder
    deterministic_base: DeterministicBase
    analytical_expectation: AnalyticalExpectation
    term_labels: tuple[str, ...]
    rng_signature: tuple
    """The operational snapshot of the RNG reference this model was prepared
    against. `run_simulation` refuses a reference that does not match it, before
    the first draw, so `rng_version` cannot describe one generator while another
    produced the numbers."""

    @property
    def drivers(self) -> tuple[PreparedSimulationDriver, ...]:
        """Canonical accumulation order: every Cost Line, then every Risk."""
        return self.cost_drivers + self.risk_drivers


# ===========================================================================
# run output
# ===========================================================================
@dataclass(frozen=True)
class ComponentDiagnostics:
    """One component's stream identity and what the run did to it.

    ORACLE DIAGNOSTICS. This is not `_SimData`, not a persisted workbook schema
    and not a retained sample. Its size is one record per component - not per
    component per iteration - and it exists so the D6-18b severity-stream
    invariance can be proved rather than asserted.
    """

    kind: str
    permanent_id: str
    role: str
    stream_index: int
    initial_state: RngState
    final_state: RngState
    uniforms_consumed: int


@dataclass(frozen=True)
class SimulationSummary:
    nominal: MeasureStatistics
    pv: MeasureStatistics
    ladder: PercentileLadder

    def measure(self, name: str) -> MeasureStatistics:
        if name == "nominal":
            return self.nominal
        if name == "pv":
            return self.pv
        raise SimOracleError(f"unknown measure {name!r}; the measures are 'nominal' and 'pv'")


@dataclass(frozen=True)
class SimulationResult:
    """The complete, immutable output of one successful run.

    The only retained stochastic arrays are `total_nominal` and `total_pv`, both
    in ORIGINAL iteration order. There is no per-driver sample matrix, no annual
    stochastic matrix and no sensitivity column: a 300-driver 100,000-iteration
    run retains 200,000 Doubles, not 30,000,000.
    """

    iterations: int
    effective_seed: int
    rng_version: int
    sim_method_version: int
    total_nominal: tuple[float, ...]
    total_pv: tuple[float, ...]
    result_digest: str
    summary: SimulationSummary
    diagnostics: tuple[ComponentDiagnostics, ...]


@dataclass(frozen=True)
class Contingency:
    """`selected Px total - deterministic base estimate A`, per measure."""

    selected_confidence_level: str
    selected_nominal: float
    selected_pv: float
    base_nominal: float
    base_pv: float
    nominal: float
    pv: float


# ===========================================================================
# iteration count pre-flight
# ===========================================================================
def business_minimum_iterations(inputs: InputContract) -> int:
    """The locked minimum, read from the input that owns it.

    Not restated here and not defaulted. The contract declares the validation as
    a whole-number `greaterThanOrEqual` rule; a contract that declared something
    else would mean the minimum had moved somewhere this function cannot see, so
    it refuses rather than guessing.
    """
    spec = inputs.inputs.get(ITERATIONS_INPUT_KEY)
    if spec is None:
        raise SimOracleError(
            f"input_contract.yaml declares no {ITERATIONS_INPUT_KEY!r} input, so the business "
            "minimum iteration count has no owner"
        )
    validation = spec.validation
    if not isinstance(validation, Mapping):
        raise SimOracleError(
            f"{ITERATIONS_INPUT_KEY}: the contract declares no validation, so the business "
            "minimum is unowned. It is not defaulted here."
        )
    kind = validation.get("kind")
    operator = validation.get("operator")
    if kind != "whole" or operator != "greaterThanOrEqual":
        raise SimOracleError(
            f"{ITERATIONS_INPUT_KEY}: this reader implements a whole-number "
            f"'greaterThanOrEqual' minimum; the contract declares kind={kind!r} "
            f"operator={operator!r}. The rule has changed shape and must not be reinterpreted."
        )
    raw = validation.get("formula1")
    try:
        minimum = int(str(raw))
    except (TypeError, ValueError):
        raise SimOracleError(
            f"{ITERATIONS_INPUT_KEY}: formula1 is {raw!r}, which is not a whole number"
        ) from None
    if minimum < 1:
        raise SimOracleError(
            f"{ITERATIONS_INPUT_KEY}: the declared minimum is {minimum}, which is not a "
            "positive iteration count"
        )
    return minimum


def validate_iterations(sim: SimContract, inputs: InputContract, iterations: Any) -> int:
    """Refuse an inadmissible iteration count BEFORE anything is allocated.

    Two owners, two different kinds of refusal, both here:

        business minimum   input_contract.yaml
        technical ceiling  sim_contract.yaml (max_iterations_representable)

    The technical ceiling is a TECHNICAL limit and says so; it is not presented
    as business validation. No smaller performance cap is invented - the Python
    oracle's elapsed runtime is not a Phase-6 gate.

    This is callable on its own, so the boundary can be proved without executing
    a million-row simulation.
    """
    if isinstance(iterations, bool) or not isinstance(iterations, int):
        raise SimOracleError(
            f"the iteration count must be a whole integer, got {iterations!r}. A bool is not "
            "an integer here even though Python says it is."
        )
    minimum = business_minimum_iterations(inputs)
    if iterations < minimum:
        raise SimOracleError(
            f"iteration count {iterations} is below the business minimum {minimum} owned by "
            f"input_contract.yaml ({ITERATIONS_INPUT_KEY})"
        )
    ceiling = sim.max_iterations_representable
    if iterations > ceiling:
        raise SimOracleError(
            f"iteration count {iterations} exceeds the TECHNICAL ceiling {ceiling}: the "
            f"_SimData sheet reserves {sim.reserved_rows_h} rows, so no further iteration can "
            "be represented. This is a technical limit, not a business rule."
        )
    return iterations


# ===========================================================================
# percentile ladder
# ===========================================================================
def _percentile_fraction(label: Any, where: str) -> float:
    if not isinstance(label, str):
        raise SimOracleError(f"{where}: a percentile label must be text, got {label!r}")
    match = _PERCENTILE_LABEL.match(label)
    if match is None:
        raise SimOracleError(
            f"{where}: {label!r} is not a percentile label of the accepted form P<number>"
        )
    number = int(match.group(1))
    if not 0 <= number <= 100:
        raise SimOracleError(f"{where}: percentile {label!r} is outside P0..P100")
    return number / 100.0


def resolve_percentile_ladder(sim: SimContract, inputs: InputContract) -> PercentileLadder:
    """Combine the contract's fixed percentiles with the selectable ladder.

    NO DUPLICATE LADDER. The selectable levels are read from the config table the
    simulation contract names as their owner, and the fixed non-selectable ones
    from the simulation contract itself. A copy in this module would be a third
    authority, free to disagree with both.

    The two sets must be disjoint and `p10_selectable` must agree with them; a
    contract that says P10 is not selectable while listing it as selectable is
    refused rather than silently resolved one way.
    """
    block = sim.raw.get("statistics")
    if not isinstance(block, Mapping):
        raise SimOracleError("sim_contract.yaml declares no 'statistics' section")

    locator = block.get("selectable_ladder_locator")
    expected_locator = f"config_tables.{CONFIDENCE_LEVELS_TABLE_KEY}"
    if locator != expected_locator:
        raise SimOracleError(
            f"sim_contract.yaml points the selectable ladder at {locator!r}; this reader "
            f"implements {expected_locator!r}. The owner has moved and must not be guessed."
        )
    owner = block.get("selectable_ladder_owner")
    if owner != "input_contract.yaml":
        raise SimOracleError(
            f"sim_contract.yaml names {owner!r} as the selectable-ladder owner; this reader "
            "reads input_contract.yaml"
        )

    table = None
    for candidate in inputs.config_tables:
        if candidate.key == CONFIDENCE_LEVELS_TABLE_KEY:
            table = candidate
            break
    if table is None:
        raise SimOracleError(
            f"input_contract.yaml declares no {CONFIDENCE_LEVELS_TABLE_KEY!r} config table, so "
            "the selectable confidence levels have no owner"
        )
    selectable = tuple(str(row[0]) for row in table.seed_rows)
    if not selectable:
        raise SimOracleError(
            f"the {CONFIDENCE_LEVELS_TABLE_KEY!r} config table declares no values"
        )

    fixed_raw = block.get("fixed_nonselectable_percentiles")
    if not isinstance(fixed_raw, Sequence) or isinstance(fixed_raw, (str, bytes)):
        raise SimOracleError(
            "sim_contract.yaml: statistics.fixed_nonselectable_percentiles must be a list"
        )
    fixed = tuple(str(label) for label in fixed_raw)

    overlap = sorted(set(fixed) & set(selectable))
    if overlap:
        raise SimOracleError(
            f"the contract states {overlap} as BOTH fixed non-selectable and selectable "
            "confidence levels. One of the two authorities is wrong; this is not resolved here."
        )

    p10_selectable = block.get("p10_selectable")
    if not isinstance(p10_selectable, bool):
        raise SimOracleError("sim_contract.yaml: statistics.p10_selectable must be a boolean")
    if p10_selectable != ("P10" in selectable):
        raise SimOracleError(
            f"sim_contract.yaml says p10_selectable={p10_selectable} while the selectable "
            f"ladder {'contains' if 'P10' in selectable else 'does not contain'} P10"
        )

    if block.get("include_all_selectable_ladder_values") is not True:
        raise SimOracleError(
            "sim_contract.yaml: statistics.include_all_selectable_ladder_values must be true; "
            "this oracle stores every selectable level and does not choose a subset"
        )

    points: list[tuple[str, float]] = []
    seen: set[str] = set()
    for label in fixed + selectable:
        if label in seen:
            raise SimOracleError(f"percentile {label!r} is declared twice")
        seen.add(label)
        points.append((label, _percentile_fraction(label, "percentile ladder")))
    points.sort(key=lambda entry: (entry[1], entry[0]))
    ordered = tuple(label for label, _ in points)

    headline_raw = block.get("headline_percentiles")
    if not isinstance(headline_raw, Sequence) or isinstance(headline_raw, (str, bytes)):
        raise SimOracleError("sim_contract.yaml: statistics.headline_percentiles must be a list")
    headline = tuple(str(label) for label in headline_raw)
    unknown = [label for label in headline if label not in seen]
    if unknown:
        raise SimOracleError(
            f"the contract nominates headline percentile(s) {unknown} that the resolved ladder "
            f"{list(ordered)} does not contain"
        )

    return PercentileLadder(
        fixed=fixed,
        selectable=selectable,
        ordered=ordered,
        headline=headline,
        points=tuple(points),
    )


# ===========================================================================
# preparation
# ===========================================================================
def effective_seed_from_nonce(reference: RngReference, nonce: int) -> int:
    """Convenience wrapper over the accepted Step-2 pure mapping.

    PURE. It persists nothing, increments nothing and allocates no `run_id`. The
    transactional nonce lifecycle - whether a failed attempt consumed a nonce,
    when the counter advances, what attempt metadata is published - belongs to
    the later reporting/state boundary and is deliberately absent here.
    """
    return reference.auto_seed_from_nonce(nonce)


def _family_of(driver: _DriverBase) -> str:
    """The sampler family for a driver, validated by the Phase-5 resolver first.

    Two vocabularies exist - the Phase-5 internal `DistributionKind` and the
    accepted display names the sampler dispatches on - and they must agree. A
    name the calculation accepts but the sampler does not is an authority
    conflict, not something to translate around.
    """
    _resolve_distribution(driver)  # refuses an unaccepted distribution, with its own message
    family = driver.distribution
    if family not in ACCEPTED_FAMILIES:
        raise SimOracleError(
            f"driver {driver.permanent_id!r}: the calculation accepts distribution "
            f"{family!r} but the sampler families are {list(ACCEPTED_FAMILIES)}. The two "
            "authorities disagree; this is not translated silently."
        )
    return family


def prepare_simulation(
    reference: RngReference,
    sim: SimContract,
    inputs: InputContract,
    model: CalculationModel,
    tolerances: Tolerances,
    *,
    effective_seed: int,
    iterations: int,
) -> tuple[PreparedSimulationModel, CalculationResult]:
    """Resolve everything, in this order, and draw nothing.

    1. refuse an inadmissible iteration count - before any allocation, any
       stream construction and any random draw;
    2. call the accepted Phase-5 `calculate` ONCE;
    3. take `Quantity`, `Probability`, `Knom`, `Kpv` and the driver identities
       from its `DriverFactors`;
    4. pair them with Min / Most Likely / Max resolved by the Phase-5 resolvers;
    5. prepare each Beta-PERT shape once;
    6. construct every component's initial stream state once;
    7. return - the loop is entered only after all of this succeeded.

    The Phase-5 `CalculationResult` is returned alongside the prepared model
    rather than stored inside it: the deterministic base A and the analytical
    expectation E are carried forward as four plain Doubles, and nothing else
    from a 300-driver audit record survives into the hot loop.
    """
    validate_iterations(sim, inputs, iterations)

    # THE SUPPLIED REFERENCE MUST BE THE ONE THESE CONTRACTS DERIVE, AND MUST BE
    # THE ACCEPTED IMPLEMENTATION. A caller-constructed RngReference is not
    # trusted: it can hold any constants at all, and a behavioural subclass can
    # hold the RIGHT constants and still generate something else entirely -
    # either way a run built on it would report the contract's `rng_version`.
    # Both are checked here, before stream construction and before any draw, and
    # the signature is snapshotted so the run boundary can check it again.
    _require_exact_reference(reference, "preparing a simulation")
    signature = rng_reference_signature(reference)
    expected = rng_reference_signature(RngReference.from_contracts(sim, inputs))
    if signature != expected:
        raise SimOracleError(
            f"the supplied RngReference is not the one {sim.source_path.name} and "
            f"{inputs.source_path.name} derive: "
            f"{_binding_mismatch(expected, signature)} differs. A simulation prepared "
            "against a reference the accepted contracts did not produce would report "
            f"rng_version {sim.rng_version} for a different generator."
        )

    ladder = resolve_percentile_ladder(sim, inputs)

    result = calculate(model, tolerances)
    factors = {
        (record.driver_kind, record.permanent_id): record for record in result.drivers
    }

    ordered_costs = canonical_order(model.cost_drivers)
    ordered_risks = canonical_order(model.risk_drivers)
    cost_ids = tuple(driver.permanent_id for driver in ordered_costs)
    risk_ids = tuple(driver.permanent_id for driver in ordered_risks)

    components = reference.components_for(cost_ids, risk_ids)
    streams: dict[tuple[str, str, str], tuple[int, RngState]] = {}
    for component, index, state in reference.component_stream_states(
        reference.fixed_seed_to_state(effective_seed), components
    ):
        streams[(component.kind, component.permanent_id, component.role)] = (index, state)

    def stream(kind: str, permanent_id: str, role: str) -> tuple[int, RngState]:
        try:
            return streams[(kind, permanent_id, role)]
        except KeyError:  # pragma: no cover - components_for builds exactly these
            raise SimOracleError(
                f"no {role} stream was assigned to {kind} {permanent_id!r}"
            ) from None

    prepared_costs: list[PreparedSimulationDriver] = []
    for driver in ordered_costs:
        record = factors.get((DriverKind.COST_LINE, driver.permanent_id))
        if record is None:  # pragma: no cover - calculate emits every driver
            raise SimOracleError(f"the calculation produced no factors for {driver.permanent_id!r}")
        family = _family_of(driver)
        minimum, most_likely, maximum = _resolve_three_point(
            driver, _resolve_distribution(driver)
        )
        index, state = stream(COST_KIND, driver.permanent_id, ROLE_VALUE)
        prepared_costs.append(
            PreparedSimulationDriver(
                permanent_id=driver.permanent_id,
                driver_kind=COST_KIND,
                distribution=family,
                minimum=minimum,
                most_likely=most_likely,
                maximum=maximum,
                quantity=record.quantity,
                probability=None,
                knom=record.knom,
                kpv=record.kpv,
                beta_shape=(
                    prepare_beta_pert(minimum, most_likely, maximum)
                    if family == FAMILY_BETA_PERT
                    else None
                ),
                value_stream_index=index,
                value_initial_state=state,
                occurrence_stream_index=None,
                occurrence_initial_state=None,
                nominal_where=f"Cost Line {driver.permanent_id!r}: nominal contribution",
                pv_where=f"Cost Line {driver.permanent_id!r}: PV contribution",
            )
        )

    prepared_risks: list[PreparedSimulationDriver] = []
    for driver in ordered_risks:
        record = factors.get((DriverKind.RISK, driver.permanent_id))
        if record is None:  # pragma: no cover - calculate emits every driver
            raise SimOracleError(f"the calculation produced no factors for {driver.permanent_id!r}")
        family = _family_of(driver)
        minimum, most_likely, maximum = _resolve_three_point(
            driver, _resolve_distribution(driver)
        )
        severity_index, severity_state = stream(
            RISK_KIND, driver.permanent_id, ROLE_SEVERITY
        )
        occurrence_index, occurrence_state = stream(
            RISK_KIND, driver.permanent_id, ROLE_OCCURRENCE
        )
        if severity_index == occurrence_index:  # pragma: no cover - assignment is injective
            raise SimOracleError(
                f"risk {driver.permanent_id!r}: occurrence and severity were assigned the same "
                "stream, which would couple the two draws"
            )
        prepared_risks.append(
            PreparedSimulationDriver(
                permanent_id=driver.permanent_id,
                driver_kind=RISK_KIND,
                distribution=family,
                minimum=minimum,
                most_likely=most_likely,
                maximum=maximum,
                quantity=None,
                probability=record.probability,
                knom=record.knom,
                kpv=record.kpv,
                beta_shape=(
                    prepare_beta_pert(minimum, most_likely, maximum)
                    if family == FAMILY_BETA_PERT
                    else None
                ),
                value_stream_index=severity_index,
                value_initial_state=severity_state,
                occurrence_stream_index=occurrence_index,
                occurrence_initial_state=occurrence_state,
                nominal_where=f"Risk {driver.permanent_id!r}: nominal contribution",
                pv_where=f"Risk {driver.permanent_id!r}: PV contribution",
            )
        )

    labels = tuple(
        f"{driver.driver_kind} {driver.permanent_id!r}"
        for driver in tuple(prepared_costs) + tuple(prepared_risks)
    )

    prepared = PreparedSimulationModel(
        cost_drivers=tuple(prepared_costs),
        risk_drivers=tuple(prepared_risks),
        iterations=iterations,
        effective_seed=effective_seed,
        base_state=reference.fixed_seed_to_state(effective_seed),
        rng_version=sim.rng_version,
        sim_method_version=sim.sim_method_version,
        ladder=ladder,
        deterministic_base=DeterministicBase(result.totals.a_nom, result.totals.a_pv),
        analytical_expectation=AnalyticalExpectation(result.totals.e_nom, result.totals.e_pv),
        term_labels=labels,
        rng_signature=signature,
    )
    return prepared, result


# ===========================================================================
# the engine
# ===========================================================================
def _sample_value(
    reference: RngReference, state: RngState, driver: PreparedSimulationDriver
):
    """One draw from a driver's own distribution, through the accepted samplers.

    Beta-PERT goes through the shape prepared once per driver, so the square root
    and the setup logarithms are not recomputed 100,000 times.
    """
    if driver.distribution == FAMILY_UNIFORM:
        return sample_uniform(reference, state, driver.minimum, driver.maximum, driver.most_likely)
    if driver.distribution == FAMILY_TRIANGULAR:
        return sample_triangular(
            reference, state, driver.minimum, driver.most_likely, driver.maximum
        )
    shape = driver.beta_shape
    if shape is None:  # pragma: no cover - prepared by prepare_simulation
        raise SimOracleError(f"driver {driver.permanent_id!r}: no prepared Beta-PERT shape")
    return sample_prepared_beta(reference, state, shape)


def run_simulation(
    reference: RngReference, prepared: PreparedSimulationModel
) -> SimulationResult:
    """The Monte Carlo loop. Deterministic, in memory, side-effect free.

    ONE ITERATION, in canonical order:

        every Cost Line, ascending Permanent ID
            sample UNIT COST from its own distribution
            nominal += unit_cost * Quantity * Knom
            pv      += unit_cost * Quantity * Kpv

        then every Risk, ascending Permanent ID
            draw occurrence from the Risk's OCCURRENCE stream (exactly one
              uniform, every iteration)
            invoke the severity sampler on the Risk's SEVERITY stream
              UNCONDITIONALLY - D6-18b - whether or not it occurred
            nominal += severity * Knom  if occurred else 0
            pv      += severity * Kpv   if occurred else 0

    Quantity is deterministic, sits OUTSIDE the distribution and is applied
    exactly once. Probability never enters `Knom` or `Kpv`; it is spent on the
    Bernoulli draw and nowhere else. PV is an independent accumulator over the
    same driver order and is never derived by discounting the sampled nominal.

    WHY SEVERITY IS SAMPLED WHEN THE RISK DID NOT OCCUR. Consumption is a
    property of the DISTRIBUTION, not of the occurrence. Sampling only on
    occurrence would make every later draw on that stream depend on the
    occurrence decisions before it, so two runs of the same model differing only
    in one Probability would produce unrelated severity sequences and could not
    be compared. Under D6-18b the severity sequence is a function of the seed and
    the distribution alone.

    NO PARTIAL RESULT. A refusal at any iteration propagates; there is no
    half-filled tuple returned as success.

    Every current RNG state lives in the lists below, one entry per component.
    No two components share one, and nothing survives the call.
    """
    if not isinstance(prepared, PreparedSimulationModel):
        raise SimOracleError(
            f"expected a PreparedSimulationModel, got {type(prepared).__name__}"
        )

    # Re-checked here, BEFORE THE FIRST DRAW - both the exact implementation and
    # the field signature. Preparation proved the reference was the contracts'
    # own; this proves the reference actually handed to the engine is still that
    # one. The component stream states were derived under the bound reference, so
    # running them through a different recurrence, a different jump ladder, a
    # different role order - or the same constants with a different
    # `next_uniform` - would produce numbers no `rng_version` describes.
    _require_exact_reference(reference, "running a simulation")
    supplied = rng_reference_signature(reference)
    if supplied != prepared.rng_signature:
        raise SimOracleError(
            "the RngReference supplied to run_simulation is not the one this model was "
            f"prepared against: {_binding_mismatch(prepared.rng_signature, supplied)} "
            f"differs. The run is refused; rng_version {prepared.rng_version} describes "
            "the bound reference and cannot be claimed for another."
        )

    iterations = prepared.iterations
    costs = prepared.cost_drivers
    risks = prepared.risk_drivers
    labels = prepared.term_labels

    cost_states: list[RngState] = [driver.value_initial_state for driver in costs]
    cost_uniforms: list[int] = [0] * len(costs)
    severity_states: list[RngState] = [driver.value_initial_state for driver in risks]
    severity_uniforms: list[int] = [0] * len(risks)
    occurrence_states: list[RngState] = [
        driver.occurrence_initial_state for driver in risks  # type: ignore[misc]
    ]
    occurrence_uniforms: list[int] = [0] * len(risks)

    total_nominal: list[float] = []
    total_pv: list[float] = []

    for index in range(1, iterations + 1):
        try:
            nominal_terms: list[float] = []
            pv_terms: list[float] = []

            for position, driver in enumerate(costs):
                drawn = _sample_value(reference, cost_states[position], driver)
                cost_states[position] = drawn.state
                cost_uniforms[position] += drawn.uniforms_consumed
                unit_cost = drawn.value
                nominal_terms.append(
                    safe_product([unit_cost, driver.quantity, driver.knom], driver.nominal_where)
                )
                pv_terms.append(
                    safe_product([unit_cost, driver.quantity, driver.kpv], driver.pv_where)
                )

            for position, driver in enumerate(risks):
                occurrence = bernoulli_occurs(
                    reference, occurrence_states[position], driver.probability
                )
                occurrence_states[position] = occurrence.state
                occurrence_uniforms[position] += occurrence.uniforms_consumed

                severity = _sample_value(reference, severity_states[position], driver)
                severity_states[position] = severity.state
                severity_uniforms[position] += severity.uniforms_consumed

                if occurrence.occurred:
                    nominal_terms.append(
                        safe_product([severity.value, driver.knom], driver.nominal_where)
                    )
                    pv_terms.append(safe_product([severity.value, driver.kpv], driver.pv_where))
                else:
                    nominal_terms.append(0.0)
                    pv_terms.append(0.0)

            total_nominal.append(
                safe_signed_sum(nominal_terms, "iteration total nominal", labels)
            )
            total_pv.append(safe_signed_sum(pv_terms, "iteration total PV", labels))
        except (CalculationRefusal, ContractError) as error:
            raise type(error)(f"iteration {index}: {error}") from error

    nominal = tuple(total_nominal)
    pv = tuple(total_pv)
    diagnostics = _diagnostics(
        costs, risks,
        cost_states, cost_uniforms,
        occurrence_states, occurrence_uniforms,
        severity_states, severity_uniforms,
    )
    return SimulationResult(
        iterations=iterations,
        effective_seed=prepared.effective_seed,
        rng_version=prepared.rng_version,
        sim_method_version=prepared.sim_method_version,
        total_nominal=nominal,
        total_pv=pv,
        result_digest=result_digest(prepared.sim_method_version, nominal, pv),
        summary=SimulationSummary(
            nominal=describe(nominal, prepared.ladder.points, "nominal"),
            pv=describe(pv, prepared.ladder.points, "PV"),
            ladder=prepared.ladder,
        ),
        diagnostics=diagnostics,
    )


def _diagnostics(
    costs: Sequence[PreparedSimulationDriver],
    risks: Sequence[PreparedSimulationDriver],
    cost_states: Sequence[RngState],
    cost_uniforms: Sequence[int],
    occurrence_states: Sequence[RngState],
    occurrence_uniforms: Sequence[int],
    severity_states: Sequence[RngState],
    severity_uniforms: Sequence[int],
) -> tuple[ComponentDiagnostics, ...]:
    """One record per component, in canonical stream order."""
    records: list[ComponentDiagnostics] = []
    for position, driver in enumerate(costs):
        records.append(
            ComponentDiagnostics(
                kind=COST_KIND,
                permanent_id=driver.permanent_id,
                role=ROLE_VALUE,
                stream_index=driver.value_stream_index,
                initial_state=driver.value_initial_state,
                final_state=cost_states[position],
                uniforms_consumed=cost_uniforms[position],
            )
        )
    for position, driver in enumerate(risks):
        records.append(
            ComponentDiagnostics(
                kind=RISK_KIND,
                permanent_id=driver.permanent_id,
                role=ROLE_OCCURRENCE,
                stream_index=driver.occurrence_stream_index,  # type: ignore[arg-type]
                initial_state=driver.occurrence_initial_state,  # type: ignore[arg-type]
                final_state=occurrence_states[position],
                uniforms_consumed=occurrence_uniforms[position],
            )
        )
        records.append(
            ComponentDiagnostics(
                kind=RISK_KIND,
                permanent_id=driver.permanent_id,
                role=ROLE_SEVERITY,
                stream_index=driver.value_stream_index,
                initial_state=driver.value_initial_state,
                final_state=severity_states[position],
                uniforms_consumed=severity_uniforms[position],
            )
        )
    return tuple(sorted(records, key=lambda record: record.stream_index))


# ===========================================================================
# result digest - D6-17
# ===========================================================================
RESULT_DIGEST_STREAM_TAG = "PCCM-RD"
RESULT_DIGEST_SECTION = "RESULT"
RESULT_DIGEST_FIELD_COUNT = 3
RESULT_DIGEST_INDEX_ORIGIN = 1


def result_digest_stream(
    sim_method_version: int,
    total_nominal: Sequence[float],
    total_pv: Sequence[float],
    decimal_separator: str = ".",
) -> str:
    """The canonical stream, exactly as the contract's grammar states it.

    ```
    stream  ::= F_S("PCCM-RD") F_I(SIM_METHOD_VERSION) section
    section ::= F_S("RESULT") F_I(record_count) record*
    record  ::= F_I(3) F_I(iteration_index) F_N(total_nominal) F_N(total_pv)
    ```

    The index is 1-based and the records are in ORIGINAL iteration order. The
    retained samples are never sorted for the digest: sorting would make two runs
    that produced the same multiset of totals in different orders indistinguishable,
    and the digest exists to identify the retained sequence, not its contents.

    The encoders come from `calc_fingerprint`. There is no competing hash here -
    this function builds a stream and hands it to the accepted primitive.

    STANDALONE. It takes arrays, not a run, so the retained Step-0 vectors can
    exercise it without a Monte Carlo engine - including the empty framing
    vector, which a real run can never produce because the business minimum is
    at least 1000 iterations.
    """
    if isinstance(sim_method_version, bool) or not isinstance(sim_method_version, int):
        raise SimOracleError(
            f"the digest version must be a whole integer, got {sim_method_version!r}"
        )
    if len(total_nominal) != len(total_pv):
        raise SimOracleError(
            f"the retained arrays disagree in length: {len(total_nominal)} nominal totals and "
            f"{len(total_pv)} PV totals. A record needs both."
        )

    records: list[Sequence[Field]] = []
    for offset, (nominal, pv) in enumerate(zip(total_nominal, total_pv)):
        records.append(
            (
                integer_field(offset + RESULT_DIGEST_INDEX_ORIGIN),
                number_field(nominal, decimal_separator),
                number_field(pv, decimal_separator),
            )
        )
    for record in records:
        if len(record) != RESULT_DIGEST_FIELD_COUNT:  # pragma: no cover - built above
            raise SimOracleError("a result-digest record must carry exactly three fields")

    return (
        text_field(RESULT_DIGEST_STREAM_TAG).encode()
        + integer_field(sim_method_version).encode()
        + encode_section(RESULT_DIGEST_SECTION, records)
    )


def result_digest(
    sim_method_version: int,
    total_nominal: Sequence[float],
    total_pv: Sequence[float],
    decimal_separator: str = ".",
) -> str:
    """The 16-character digest of the retained iteration totals. Equality exact."""
    return fingerprint(
        result_digest_stream(sim_method_version, total_nominal, total_pv, decimal_separator)
    )


def validate_result_digest_contract(sim: SimContract) -> None:
    """Check the contract still describes the grammar this module implements.

    The constants above are not a second authority; this is the assertion that
    they are still a faithful copy. A contract that changed the tag, the section
    name, the field count or the index origin without this module changing would
    otherwise be silently ignored.
    """
    block = sim.raw.get("result_digest")
    if not isinstance(block, Mapping):
        raise SimOracleError("sim_contract.yaml declares no 'result_digest' section")
    expected = {
        "stream_tag": RESULT_DIGEST_STREAM_TAG,
        "section_name": RESULT_DIGEST_SECTION,
        "record_field_count": RESULT_DIGEST_FIELD_COUNT,
        "iteration_index_origin": RESULT_DIGEST_INDEX_ORIGIN,
        "version_field_source": "sim_method_version",
        "samples_sorted_for_digest": False,
        "equality": "exact",
        "order_source": "persisted_iteration_order",
    }
    for key, want in expected.items():
        got = block.get(key)
        if got != want:
            raise SimOracleError(
                f"sim_contract.yaml: result_digest.{key} is {got!r}; this implementation "
                f"encodes {want!r}"
            )
    if list(block.get("record_fields") or ()) != ["iteration_index", "total_nominal", "total_pv"]:
        raise SimOracleError(
            f"sim_contract.yaml: result_digest.record_fields is "
            f"{block.get('record_fields')!r}; this implementation encodes iteration_index, "
            "total_nominal, total_pv in that order"
        )
    if list(block.get("field_types") or ()) != ["F_I", "F_N", "F_N"]:
        raise SimOracleError(
            f"sim_contract.yaml: result_digest.field_types is {block.get('field_types')!r}; "
            "this implementation encodes F_I, F_N, F_N"
        )


# ===========================================================================
# contingency - reporting only
# ===========================================================================
def contingency_at(
    summary: SimulationSummary,
    selected_confidence_level: str,
    deterministic_base: DeterministicBase,
) -> Contingency:
    """`selected Px total - A`, for nominal and PV, from statistics already computed.

    REPORTING ONLY. Changing the Selected Confidence Level reruns no RNG, alters
    no retained sample, no `result_digest`, no simulation mean and no stored
    percentile: every level in the ladder was computed during the run, and this
    function chooses among them.

    P10 IS NOT A SELECTOR. It is stored and reported, and the contract makes it
    non-selectable; asking for it here is refused rather than quietly served.

    THE BASELINE IS THE PHASE-5 DETERMINISTIC BASE ESTIMATE A. Not the simulation
    mean, not the analytical expected total, not `A + EMV`. The result may be
    negative when the selected percentile falls below A, and it is NOT clamped -
    a negative contingency is a real statement about the model.
    """
    if not isinstance(summary, SimulationSummary):
        raise SimOracleError(f"expected a SimulationSummary, got {type(summary).__name__}")
    if not isinstance(deterministic_base, DeterministicBase):
        raise SimOracleError(
            f"expected a DeterministicBase, got {type(deterministic_base).__name__}"
        )
    ladder = summary.ladder
    if selected_confidence_level not in ladder.selectable:
        if selected_confidence_level in ladder.ordered:
            raise SimOracleError(
                f"{selected_confidence_level!r} is reported but is not selectable; the "
                f"selectable levels are {list(ladder.selectable)}"
            )
        raise SimOracleError(
            f"{selected_confidence_level!r} is not an accepted confidence level; the "
            f"selectable levels are {list(ladder.selectable)}"
        )

    selected_nominal = summary.nominal.percentile(selected_confidence_level)
    selected_pv = summary.pv.percentile(selected_confidence_level)
    return Contingency(
        selected_confidence_level=selected_confidence_level,
        selected_nominal=selected_nominal,
        selected_pv=selected_pv,
        base_nominal=deterministic_base.nominal,
        base_pv=deterministic_base.pv,
        nominal=safe_subtract(
            selected_nominal, deterministic_base.nominal, "contingency nominal"
        ),
        pv=safe_subtract(selected_pv, deterministic_base.pv, "contingency PV"),
    )


def deterministic_base_of(result: CalculationResult) -> DeterministicBase:
    """The accepted Phase-5 `A` totals, as the contingency baseline value type."""
    return DeterministicBase(result.totals.a_nom, result.totals.a_pv)
