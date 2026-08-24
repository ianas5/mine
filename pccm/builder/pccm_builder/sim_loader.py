"""Load and validate the PCCM simulation contract.

`spec/sim_contract.yaml` is the sixth authority. It owns the simulation
EXECUTION SEMANTICS settled by Phase-6 Step 0 and nothing else. This loader is
what makes "and nothing else" enforceable rather than aspirational.

Four classes of assertion live here.

**Authority conformance.** Every value Step 0 settled - the MRG constants, the
jump matrices, the Cheng literals and expression order, the version ownership
table, the digest framing - is a LOCKED CONSTANT in this module and the contract
is checked against it. The contract ENCODES the accepted authority; it does not
get to choose it. A one-token mutation fails.

**Exclusion.** Three things must be absent, and absence is checked by scanning,
not by trusting the schema: the admissible seed RANGE (owned by
`input_contract.yaml`), any business iteration MAXIMUM, and any oracle
comparison TOLERANCE. Each is a rule that belongs somewhere else, and a second
copy here is exactly the drift this architecture exists to prevent.

**Derivation.** The technical storage ceiling is RECOMPUTED from the declared
`_SimData` layout and compared with the declared constant. D6-08 therefore
cannot be satisfied by a free literal that happens to look plausible.

**Boundary.** Every borrowed value is declared as a reference whose locator
resolves against the owning file, exactly as the calculation contract does.

Fails loudly; never repairs. Nothing here simulates: no state advances, no
uniform is generated, no jump is executed, no variate is sampled.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from .contract_loader import ContractError, InputContract
from .driver_loader import DriverContract
from .spec_loader import WorkbookSpec
from .structure_loader import StructureContract

MAX_EXCEL_ROWS = 1048576
"""Excel's worksheet row capacity. A physical limit, not a business rule."""


class SimContractError(ContractError):
    """Raised when the simulation contract is invalid.

    Subclasses ContractError so the build entry point reports every specification
    failure the same way.
    """


# ---------------------------------------------------------------------------
# The parser boundary - duplicate mapping keys are a contract defect
# ---------------------------------------------------------------------------
class _DuplicateMappingKey(Exception):
    """Internal signal from the strict loader; converted to SimContractError."""

    def __init__(self, key: Any, first: Any, second: Any) -> None:
        super().__init__(key)
        self.key = key
        self.first = first
        self.second = second


class _StrictYamlLoader(yaml.SafeLoader):
    """`SafeLoader` that refuses a mapping key declared more than once.

    The same parser-boundary guard the calculation contract uses, for the same
    reason: PyYAML resolves duplicates silently and last-wins, so every rule this
    file locks would otherwise be defeatable by writing the field twice. A reader
    would see one value and the validator another.
    """

    def construct_mapping(self, node: Any, deep: bool = False) -> dict[Any, Any]:
        seen: dict[Any, Any] = {}
        for key_node, _ in node.value:
            key = self.construct_object(key_node, deep=deep)
            try:
                duplicate = key in seen
            except TypeError:
                continue
            if duplicate:
                raise _DuplicateMappingKey(key, seen[key], key_node.start_mark)
            seen[key] = key_node.start_mark
        return super().construct_mapping(node, deep=deep)


def _strict_safe_load(text: str, path: Path) -> Any:
    try:
        return yaml.load(text, _StrictYamlLoader)  # noqa: S506 - strict SafeLoader subclass
    except _DuplicateMappingKey as duplicate:
        raise SimContractError(
            f"{path}: duplicate key {duplicate.key!r} in the same mapping - first declared at "
            f"line {duplicate.first.line + 1}, column {duplicate.first.column + 1}, again at "
            f"line {duplicate.second.line + 1}, column {duplicate.second.column + 1}. "
            "A contract must not contain two competing values for one field."
        ) from duplicate
    except yaml.YAMLError as error:
        raise SimContractError(f"{path}: is not valid YAML: {error}") from error


# ---------------------------------------------------------------------------
# The accepted Step-0 authority - LOCKED HERE, not in the contract
# ---------------------------------------------------------------------------
LOCKED_SIM_CONTRACT_VERSION = "1.0.0"
"""The contract-document format this loader implements.

A different version domain from RNG_VERSION, SIM_METHOD_VERSION and FP_VERSION.
All four move independently and none implies another."""

LOCKED_RNG_FAMILY = "MRG32k3a"
LOCKED_RNG_CONSTANTS = {
    "m1": 4294967087,
    "m2": 4294944443,
    "a12": 1403580,
    "a13n": 810728,
    "a21": 527612,
    "a23n": 1370589,
    "norm": 2.328306549295727688e-10,
}
LOCKED_STATE_ORDER = ("s10", "s11", "s12", "s20", "s21", "s22")

LOCKED_A1_P127 = (
    (1230515664, 986791581, 1988835001),
    (3580155704, 1230515664, 226153695),
    (949770784, 3580155704, 2427906178),
)
LOCKED_A2_P127 = (
    (2093834863, 32183930, 2824425944),
    (1022607788, 1464411153, 32183930),
    (1610723613, 277697599, 1464411153),
)
LOCKED_JUMP_EXPONENT = 127
LOCKED_JUMP_H = 1 << 17

LOCKED_AUTO_MODULUS = 2147483647
LOCKED_AUTO_MULTIPLIER = 48271
LOCKED_AUTO_PERIOD = 2147483646
LOCKED_NONCE_EXHAUSTED = 2147483646
LOCKED_NONCE_LIFECYCLE = (
    "validate_pre_allocation_prerequisites",
    "read_current_auto_nonce",
    "derive_effective_seed",
    "persist_auto_nonce_plus_one",
    "begin_sampling",
)

LOCKED_COMPONENT_KINDS = ("COST_SAMPLE", "RISK_OCCURRENCE", "RISK_SEVERITY")
LOCKED_SORT_KEYS = ("component_kind", "permanent_id", "role")
LOCKED_ID_COMPARISON = "ordinal_utf16_code_units"

LOCKED_FAMILIES = ("Uniform", "Triangular", "Beta-PERT")
LOCKED_PERT_LAMBDA = 4
LOCKED_DISPATCH_RULE = "min(alpha, beta) > 1 -> BB; otherwise BC"
LOCKED_DISPATCH_EQUALITY_OWNER = "BC"

LOCKED_LOGIT_FORM = "log(u1 / (1.0 - u1))"
LOCKED_REJECTED_LOGIT_FORM = "log(u1) - log1p(-u1)"
LOCKED_BB_LITERALS = ("1.3862944", "2.609438", "5.0", "2.0", "1.0")
LOCKED_BC_LITERALS = ("0.0138889", "0.0416667", "0.777778", "1.3862944", "0.25", "0.5")
LOCKED_BB_PER_DRIVER = (
    "alpha = a + b",
    "beta = sqrt((alpha - 2) / (2 * a * b - alpha))",
    "gamma = a + 1 / beta",
)
LOCKED_BB_PER_ATTEMPT = (
    "u1 = next_u(); u2 = next_u()",
    "vlog = log(u1 / (1 - u1))",
    "v = beta * vlog",
    "w = a * exp(v)",
    "z = u1 * u1 * u2",
    "rr = gamma * v - 1.3862944",
    "s = a + rr - w",
    "accept if s + 2.609438 >= 5.0 * z",
    "else t = log(z); accept if s >= t",
    "else accept if rr + alpha * log(alpha / (b + w)) >= t",
    "else reject and retry",
)
LOCKED_BB_ORIENTATION = {"a": "min(alpha0, beta0)", "b": "max(alpha0, beta0)"}
LOCKED_BC_ORIENTATION = {"a": "max(alpha0, beta0)", "b": "min(alpha0, beta0)"}
LOCKED_UNIFORMS_PER_ATTEMPT = 2

LOCKED_DIGEST_TAG = "PCCM-RD"
LOCKED_DIGEST_SECTION = "RESULT"
LOCKED_DIGEST_RECORD_FIELDS = ("iteration_index", "total_nominal", "total_pv")
LOCKED_DIGEST_FIELD_TYPES = ("F_I", "F_N", "F_N")
LOCKED_VERSION_FIELD_SOURCE = "sim_method_version"

LOCKED_SECTION_ORDER = ("HEADER", "COST", "RISK", "SIM")
LOCKED_ANALYTICAL_PREFIX = ("HEADER", "COST", "RISK")
LOCKED_SIM_FIELDS = (
    "iterations",
    "seed_mode",
    "supplied_seed",
    "rng_version",
    "sim_method_version",
)
LOCKED_SIM_EXCLUDED = ("effective_seed", "auto_nonce", "run_id", "selected_confidence_level")

LOCKED_SIM_STATES = ("CURRENT", "STALE", "INVALID")
LOCKED_ATTEMPT_RESULTS = ("NONE", "SUCCESS", "REFUSED", "FAILED")
LOCKED_SEED_MODES = ("AUTO", "FIXED")

LOCKED_PERCENTILE_METHOD = "hyndman_fan_type_7"
LOCKED_SD_DIVISOR = "n_minus_1"
LOCKED_CONTINGENCY_BASELINE = "deterministic_base_estimate_a"
LOCKED_FORBIDDEN_BASELINES = ("simulation_mean", "analytical_expected_total", "a_plus_emv")

LOCKED_RUN_ID_MAX = 2147483647

LOCKED_RNG_VERSION_BUMPS = (
    "mrg32k3a_constants_or_uniform_combination",
    "scalar_seed_to_state_mapping",
    "auto_seed_nonce_mapping",
    "component_stream_assignment",
    "jump_matrices_exponent_or_arithmetic_semantics",
)
LOCKED_SIM_METHOD_VERSION_BUMPS = (
    "cheng_formulation_literals_or_expression_order",
    "risk_severity_invocation_semantics",
    "degenerate_sampling_rule",
    "canonical_accumulation_order",
    "percentile_or_statistical_method",
    "result_digest_framing_or_hash_stream",
)

LOCKED_SIM_DATA_SHEET = "_SimData"
LOCKED_ITERATION_COLUMNS = ("iteration_index", "total_nominal", "total_pv")
LOCKED_RUN_IDENTITY = (
    ("last_successful_stamp", 8, "snapshot", "timestamp"),
    ("run_id", 9, "snapshot", "integer"),
    ("request_fingerprint", 10, "snapshot", "text"),
    ("result_digest", 11, "snapshot", "text"),
    ("seed_mode", 12, "snapshot", "enum"),
    ("supplied_seed", 13, "snapshot", "integer"),
    ("effective_seed", 14, "snapshot", "integer"),
    ("consumed_auto_nonce", 15, "snapshot", "integer"),
    ("iterations_run", 16, "snapshot", "integer"),
    ("rng_version", 17, "snapshot", "integer"),
    ("sim_method_version", 18, "snapshot", "integer"),
    ("model_version", 19, "snapshot", "text"),
    ("applied_timeline", 20, "snapshot", "text"),
    ("next_auto_nonce", 21, "counter", "integer"),
    ("last_run_id", 22, "counter", "integer"),
    ("last_attempt_result", 23, "attempt", "enum"),
    ("last_attempt_detail", 24, "attempt", "text"),
    ("last_attempt_seed_mode", 25, "attempt", "enum"),
    ("last_attempt_effective_seed", 26, "attempt", "integer"),
    ("last_attempt_auto_nonce", 27, "attempt", "integer"),
    ("simulation_status", 28, "derived", "enum"),
    ("status_evaluated_at", 29, "derived", "timestamp"),
)
"""The run-identity block is EXACT authority: key, row, group and type, in order.

"Contains these required fields" was not enough - an invented field could be
appended on the next free row and the loader accepted it. A snapshot layout that
can grow by accident is one whose meaning drifts silently."""

LOCKED_SIM_STATE_RULES = (
    (1, "current_prerequisites_do_not_resolve", "INVALID"),
    (2, "no_successful_snapshot_exists", None),
    (3, "request_fingerprint_equals_stored_successful", "CURRENT"),
    (4, "request_fingerprint_differs_from_stored_successful", "STALE"),
)
"""The corrected derivation - ordered, total, and blind to the attempt history."""

LOCKED_SIM_DATA_EXCLUDED = (
    "per_driver_samples",
    "annual_stochastic_samples",
    "sensitivity_data",
)

REQUIRED_SECTIONS = (
    "sim_contract_version",
    "versions",
    "rng",
    "seeding",
    "components",
    "stream_assignment",
    "jump",
    "distributions",
    "cheng",
    "risk",
    "accumulation",
    "contribution",
    "kernel",
    "numerical_domain",
    "dependence",
    "publication",
    "command_surface",
    "interruption",
    "request_fingerprint",
    "result_digest",
    "iterations",
    "sim_data",
    "label_sets",
    "sim_state",
    "prerequisite",
    "run_id",
    "statistics",
    "contingency",
    "results_minimum",
    "authority_references",
)

REQUIRED_AUTHORITY_REFERENCES = (
    ("Random Seed admissible domain", "input_contract.yaml", "inputs.random_seed"),
    (
        "Monte Carlo iterations business minimum",
        "input_contract.yaml",
        "inputs.monte_carlo_iterations",
    ),
    (
        "selectable confidence level ladder",
        "input_contract.yaml",
        "config_tables.confidence_levels",
    ),
    ("distribution master list", "input_contract.yaml", "config_tables.distributions"),
    ("Cost Line and Risk Register input schemas", "driver_contract.yaml", "registers"),
    (
        "permanent-ID prefixes, patterns and counter rules",
        "structure_contract.yaml",
        "identity.counters",
    ),
    ("applied timeline and structural limits", "structure_contract.yaml", "timeline"),
    ("fingerprint algorithm version", "calc_contract.yaml", "fingerprint.version"),
    (
        "deterministic base estimate and analytical totals",
        "calc_contract.yaml",
        "scalar_blocks.calc_totals",
    ),
    ("_SimData sheet declaration and visibility", "workbook.yaml", "sheets._SimData"),
    ("Results sheet placeholder sections", "workbook.yaml", "sheets.Results"),
    ("model version", "workbook.yaml", "model.model_version"),
)
"""(concept, owner, locator) - the COMPLETE required set, exactly.

Checking only that the references present resolve is not enough: a reference can
be DELETED, and a deleted boundary is silently unowned."""

ALLOWED_AUTHORITY_REFERENCE_KEYS = frozenset({"concept", "owner", "locator"})

# ---------------------------------------------------------------------------
# Exclusion scanning
# ---------------------------------------------------------------------------
SEED_RANGE_TOKENS = (
    "2147483646",
    "seed_min",
    "seed_max",
    "seed_minimum",
    "seed_maximum",
    "admissible_seed",
)
"""Tokens that would mean the seed RANGE had been copied here.

`2147483646` legitimately appears as the AUTO cycle PERIOD and as the nonce
exhaustion point, which are different facts about the same number. The scan is
therefore keyed on WHERE the token appears, not merely that it appears: see
`_forbid_seed_range`."""

SEED_RANGE_ALLOWED_LOCATIONS = (
    ("seeding", "auto", "period"),
    ("seeding", "nonce_lifecycle", "exhausted_value"),
)

TOLERANCE_TOKENS = (
    "tolerance",
    "ulp",
    "rel_tol",
    "abs_tol",
    "atol",
    "rtol",
    "epsilon",
    "comparison_tolerance",
)

WILDCARD_TOKENS = ("*", "all", "any", "all_modules", "*.bas", "**")


# ---------------------------------------------------------------------------
# CLOSED-WORLD SCHEMA
#
# Every mapping in the contract has an explicit allowed-key set, and a mapping
# reached by a path that is not listed here is itself a defect. An unknown key
# was previously ACCEPTED - `root.future_semantic`, `rng.future_semantic` and
# four more were demonstrated - which meant a semantic could be added to the
# authority document and go unread by the validator that is supposed to enforce
# it. Silence is the worst possible answer there: the field looks authoritative
# and governs nothing.
#
# `[]` in a path denotes a list element, so `components.kinds[]` is the shape of
# every record in that list. Membership of the LISTS themselves is locked
# separately by the section validators - closing the mappings alone would still
# let an invented record be appended.
# ---------------------------------------------------------------------------
CLOSED_KEYS: dict[str, frozenset[str]] = {
    '': frozenset({
        'accumulation', 'authority_references', 'cheng', 'command_surface', 'components',
        'contingency', 'contribution', 'dependence', 'distributions', 'interruption',
        'iterations', 'jump', 'kernel', 'label_sets', 'numerical_domain', 'prerequisite',
        'publication', 'request_fingerprint', 'result_digest', 'results_minimum', 'risk',
        'rng', 'run_id', 'seeding', 'sim_contract_version', 'sim_data', 'sim_state',
        'statistics', 'stream_assignment', 'versions'
    }),
    'accumulation': frozenset({
        'accumulators', 'accumulators_share_driver_order', 'driver_kind_order',
        'permanent_id_comparison', 'physical_row_order_permitted', 'within_kind_order'
    }),
    'authority_references[]': frozenset({'concept', 'locator', 'owner'}),
    'cheng': frozenset({
        'algebraic_simplification_permitted', 'bb', 'bc', 'conformance_vectors',
        'literal_effect', 'literals_are_literal', 'logit_form',
        'logit_form_rejected_alternative', 'source_binding',
        'uniforms_per_non_degenerate_proposal_attempt'
    }),
    'cheng.bb': frozenset({
        'acceptance_operator', 'applies_when', 'literals', 'orientation', 'per_attempt',
        'per_driver', 'return'
    }),
    'cheng.bb.orientation': frozenset({'a', 'b'}),
    'cheng.bc': frozenset({
        'acceptance_operator', 'applies_when', 'literals', 'orientation', 'per_attempt',
        'per_driver', 'return'
    }),
    'cheng.bc.orientation': frozenset({'a', 'b'}),
    'cheng.conformance_vectors': frozenset({'evidence_file', 'role', 'runtime_lookup_table'}),
    'cheng.literal_effect': frozenset({'logit_form_affects', 'squeeze_literals_affect'}),
    'cheng.source_binding': frozenset({'evidence_file', 'functions_sha256'}),
    'command_surface': frozenset({
        'automation_endpoint', 'msgbox_introduced_by_phase_6', 'read_accessor_names_settled',
        'ribbon_introduced_by_phase_6', 'user_facing_run_button_in_phase_6',
        'userform_introduced_by_phase_6'
    }),
    'components': frozenset({'count_rule', 'kinds'}),
    'components.kinds[]': frozenset({'driver_kind', 'key', 'per_driver', 'role'}),
    'contingency': frozenset({
        'baseline', 'forbidden_baselines', 'formula', 'measures',
        'workbook_recommends_a_confidence_level'
    }),
    'contribution': frozenset({
        'cost_line', 'iteration_total', 'pv_derived_from_nominal', 'risk'
    }),
    'contribution.cost_line': frozenset({
        'nominal', 'probability_applies', 'pv', 'quantity_applications',
        'quantity_inside_distribution', 'quantity_is_deterministic', 'sampled_from',
        'sampled_quantity', 'total_cost_uncertainty_sampled'
    }),
    'contribution.iteration_total': frozenset({'measures_independent', 'order_source', 'rule'}),
    'contribution.risk': frozenset({
        'nominal_when_not_occurred', 'nominal_when_occurred', 'occurred',
        'occurrence_and_severity_share_a_stream', 'probability_folded_into_k_factors',
        'pv_when_not_occurred', 'pv_when_occurred', 'quantity_applies', 'severity_source'
    }),
    'dependence': frozenset({
        'authority', 'copula_supported', 'correlation_matrix_supported',
        'inter_driver_dependence', 'shared_or_hidden_dependence_permitted'
    }),
    'distributions': frozenset({'beta_pert', 'degenerate', 'families', 'triangular', 'uniform'}),
    'distributions.beta_pert': frozenset({
        'alpha', 'alpha_plus_beta', 'beta', 'conditioning_scale', 'dispatch', 'lambda',
        'normalised_formulation_required', 'rescale', 'rescale_formulation', 'shape_lower',
        'shape_ratio', 'shape_upper'
    }),
    'distributions.beta_pert.dispatch': frozenset({
        'comparison_operator', 'equality_belongs_to', 'rule'
    }),
    'distributions.degenerate': frozenset({
        'applies_to_all_families', 'condition', 'detected_before_dispatch',
        'detected_before_parameterisation', 'returns', 'sampler_entered',
        'stream_state_changed', 'uniforms_consumed'
    }),
    'distributions.triangular': frozenset({
        'boundary_cases', 'branch_point', 'conditioning_scale', 'lower_branch', 'method',
        'normalised_formulation_required', 'rng_endpoints_open',
        'uniforms_per_non_degenerate_sample', 'upper_branch'
    }),
    'distributions.triangular.boundary_cases': frozenset({'m_equals_a', 'm_equals_b'}),
    'distributions.uniform': frozenset({
        'formulation', 'most_likely_used', 'transform', 'uniforms_per_non_degenerate_sample'
    }),
    'interruption': frozenset({'user_cancellation_supported_in_phase_6'}),
    'iterations': frozenset({'business_maximum', 'business_minimum_owner', 'technical_ceiling'}),
    'iterations.technical_ceiling': frozenset({
        'consumes_auto_nonce', 'max_excel_rows', 'max_iterations_representable',
        'presented_as_business_validation', 'refusal_kind', 'refusal_precedes',
        'reserved_rows_h'
    }),
    'jump': frozenset({
        'a1_p127', 'a1_p127_sha256', 'a2_p127', 'a2_p127_sha256', 'decomposition',
        'decomposition_h', 'naive_floating_matrix_product_permitted',
        'stream_spacing_exponent', 'substream_spacing_exponent', 'substreams_used_in_phase_6'
    }),
    'kernel': frozenset({
        'application_object_access_inside_iteration_loop',
        'com_round_trip_inside_iteration_loop', 'inputs_resolved_once_before_simulation',
        'listobject_access_inside_iteration_loop', 'operates_on_resolved_in_memory_structures',
        'range_access_inside_iteration_loop', 'recomputes_worksheet_fx_inside_loop',
        'recomputes_worksheet_inflation_inside_loop',
        'recomputes_worksheet_profiles_inside_loop', 'resolved_before_loop',
        'thisworkbook_or_activeworkbook_access_inside_iteration_loop',
        'worksheet_access_inside_iteration_loop'
    }),
    'label_sets': frozenset({'attempt_result', 'seed_mode', 'sim_state'}),
    'numerical_domain': frozenset({
        'disciplines', 'magnitude_restriction', 'narrower_than_phase5',
        'negative_values_legal', 'positivity_rule', 'refusal_when_no_valid_double_result',
        'representable_result_refused_for_naive_intermediate_overflow',
        'silent_non_finite_result_permitted', 'supports_crossing_zero_legal'
    }),
    'numerical_domain.disciplines': frozenset({
        'accumulation', 'contingency_subtraction', 'driver_contribution',
        'percentile_interpolation', 'statistics'
    }),
    'prerequisite': frozenset({
        'phase5_analytical_state_required', 'phase6_may_call_pccm_calculate',
        'silent_recalculation_permitted'
    }),
    'publication': frozenset({
        'commit_last', 'partial_new_distribution_published_on_refusal_or_failure',
        'persisted_source_of_truth', 'prior_successful_publication_survives',
        'publish_only_after_simulation_and_statistics_complete', 'results_derives_from',
        'results_recomputes_monte_carlo'
    }),
    'request_fingerprint': frozenset({
        'analytical_prefix', 'auto_blank_seed_remains_recomputable',
        'existing_sections_modified', 'extension_semantics', 'section_order', 'sim_section'
    }),
    'request_fingerprint.sim_section': frozenset({
        'analytical_fingerprint_hashed_as_a_field', 'excluded_fields', 'fields', 'name',
        'supplied_seed_present_only_when'
    }),
    'result_digest': frozenset({
        'equality', 'field_types', 'grammar', 'iteration_index_origin', 'order_source',
        'record_field_count', 'record_fields', 'samples_sorted_for_digest', 'section_name',
        'stream_tag', 'version_field_source'
    }),
    'result_digest.grammar': frozenset({'record', 'section', 'stream'}),
    'results_minimum': frozenset({'annual_simulated_samples_contracted', 'deferred', 'sections'}),
    'risk': frozenset({
        'occurrence', 'probability_folded_into_knom', 'probability_folded_into_kpv',
        'severity'
    }),
    'risk.occurrence': frozenset({
        'comparison_operator', 'probability_one_always_occurs',
        'probability_zero_never_occurs', 'rule', 'uniforms_per_risk_per_iteration'
    }),
    'risk.severity': frozenset({
        'degenerate_consumption', 'degenerate_stream_state_changed', 'invocation_policy',
        'non_degenerate_consumption', 'sampler_invoked_every_risk_iteration',
        'value_used_only_when_occurred'
    }),
    'rng': frozenset({
        'arithmetic', 'combination', 'constants', 'family', 'output_domain', 'recurrence',
        'state'
    }),
    'rng.arithmetic': frozenset({
        'implementation_requirement', 'modulo_semantics', 'naive_floating_modulo_permitted'
    }),
    'rng.combination': frozenset({'comparison_operator', 'rule'}),
    'rng.constants': frozenset({'a12', 'a13n', 'a21', 'a23n', 'm1', 'm2', 'norm'}),
    'rng.output_domain': frozenset({'lower', 'lower_inclusive', 'upper', 'upper_inclusive'}),
    'rng.recurrence': frozenset({'advance', 'p1', 'p2'}),
    'rng.state': frozenset({
        'matrix_operand_orientation', 'order', 'orientation',
        'reversal_required_at_matrix_boundary', 'words'
    }),
    'run_id': frozenset({
        'allocated_on', 'failure_consumes', 'first_successful_value', 'independent_of',
        'initial', 'maximum', 'on_exhaustion', 'persisted', 'reuse_permitted',
        'wrap_permitted'
    }),
    'seeding': frozenset({
        'auto', 'blank_input_means', 'modes', 'nonce_lifecycle', 'populated_input_means',
        'scalar_to_state'
    }),
    'seeding.auto': frozenset({
        'cross_workbook_uniqueness_claimed', 'freshness_scope',
        'implementation_complexity_requirement', 'mapping', 'mapping_kind', 'modulus',
        'multiplier', 'period', 'stepped_multiplication_is_the_authority',
        'timestamp_derived_uniqueness_permitted'
    }),
    'seeding.nonce_lifecycle': frozenset({
        'attempt_metadata_preserves', 'exhausted_value',
        'failure_after_allocation_consumes_nonce', 'failure_before_allocation_consumes_nonce',
        'first_valid_allocation', 'initial', 'last_valid_allocation', 'meaning',
        'on_exhaustion', 'order', 'prior_successful_publication_untouched', 'reuse_permitted',
        'wrap_permitted'
    }),
    'seeding.scalar_to_state': frozenset({
        'alternate_expansion_permitted', 'expansion', 'mixer', 'rule'
    }),
    'sim_data': frozenset({
        'excluded', 'iteration_records', 'required_visibility', 'reserved_rows',
        'run_identity', 'sheet'
    }),
    'sim_data.iteration_records': frozenset({
        'columns', 'first_iteration_row', 'footer_rows', 'header_row', 'order', 'sorted'
    }),
    'sim_data.iteration_records.columns[]': frozenset({'column', 'header', 'key', 'value_type'}),
    'sim_data.reserved_rows[]': frozenset({'purpose', 'rows'}),
    'sim_data.run_identity': frozenset({
        'fields', 'first_row', 'label_column', 'last_row', 'note_column', 'value_column'
    }),
    'sim_data.run_identity.fields[]': frozenset({
        'enum', 'group', 'initial', 'key', 'label', 'row', 'value_type'
    }),
    'sim_state': frozenset({
        'attempt_axis_is_orthogonal', 'attempt_result_participates_in_derivation',
        'definitions', 'derivation', 'no_success_valid_status', 'on_failure', 'states',
        'status_evaluated_at_may_be_populated_while_status_is_blank'
    }),
    'sim_state.definitions': frozenset({'CURRENT', 'INVALID', 'STALE'}),
    'sim_state.derivation': frozenset({'ordered', 'rules'}),
    'sim_state.derivation.rules[]': frozenset({'condition', 'order', 'status'}),
    'sim_state.on_failure': frozenset({
        'attempt_metadata_updated', 'partial_distribution_published',
        'prior_results_publication_preserved', 'prior_sim_data_preserved'
    }),
    'statistics': frozenset({
        'fixed_nonselectable_percentiles', 'headline_percentiles',
        'include_all_selectable_ladder_values', 'mean', 'measures', 'moments_and_extremes',
        'p10_selectable', 'percentile', 'selectable_ladder_locator', 'selectable_ladder_owner',
        'selected_confidence_level', 'sorting', 'standard_deviation'
    }),
    'statistics.mean': frozenset({'method', 'scale_safe_required'}),
    'statistics.percentile': frozenset({'formula', 'interpolation', 'method'}),
    'statistics.percentile.formula': frozenset({'f', 'h', 'hi', 'lo', 'value'}),
    'statistics.selected_confidence_level': frozenset({
        'affects_staleness', 'enters_request_fingerprint', 'enters_simulation_execution',
        'role'
    }),
    'statistics.standard_deviation': frozenset({
        'divisor', 'method', 'naive_sum_of_squares_permitted'
    }),
    'stream_assignment': frozenset({
        'accepted_consequence', 'case_insensitive_order_permitted', 'index_origin',
        'index_rule', 'locale_order_permitted', 'numeric_suffix_interpretation_permitted',
        'permanent_id_comparison', 'physical_row_order_permitted', 'policy', 'sort_keys'
    }),
    'versions': frozenset({
        'bump_ownership', 'result_digest_version_source', 'rng_version', 'sim_method_version'
    }),
    'versions.bump_ownership': frozenset({'rng_version', 'sim_method_version'}),
}


def _check_closed_world(node: Any, path: str, source: Path) -> None:
    """Refuse an unknown key, and refuse a mapping at an unknown path."""
    if isinstance(node, dict):
        allowed = CLOSED_KEYS.get(path)
        if allowed is None:
            raise SimContractError(
                f"{source}: a mapping appears at {path or 'the document root'!r}, which the "
                "closed-world schema does not describe. Every mapping in this contract must "
                "have a declared shape."
            )
        unknown = sorted(set(node) - allowed)
        if unknown:
            raise SimContractError(
                f"{source}: {path or 'root'} declares unknown key(s) {unknown}. This contract is "
                "CLOSED: a key the validator does not know is a semantic nobody enforces, and a "
                "field that looks authoritative while governing nothing is worse than no field."
            )
        for key, value in node.items():
            _check_closed_world(value, f"{path}.{key}" if path else key, source)
    elif isinstance(node, list):
        for value in node:
            _check_closed_world(value, f"{path}[]", source)


# ---------------------------------------------------------------------------
# The parsed contract
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class AuthorityReference:
    concept: str
    owner: str
    locator: str


@dataclass(frozen=True)
class SimDataLayout:
    """The future `_SimData` machine layout, and the ceiling it determines."""

    sheet: str
    required_visibility: str
    reserved_rows: tuple[tuple[int, int, str], ...]
    header_row: int
    first_iteration_row: int
    footer_rows: int
    reserved_row_count: int

    @property
    def max_iterations_representable(self) -> int:
        return MAX_EXCEL_ROWS - self.reserved_row_count


@dataclass(frozen=True)
class SimContract:
    """The parsed, validated simulation contract."""

    version: str
    rng_version: int
    sim_method_version: int
    layout: SimDataLayout
    authority_references: tuple[AuthorityReference, ...]
    raw: dict[str, Any]
    source_path: Path

    @property
    def max_iterations_representable(self) -> int:
        return self.layout.max_iterations_representable

    @property
    def reserved_rows_h(self) -> int:
        return self.layout.reserved_row_count


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def load_sim_contract(path: Path) -> SimContract:
    """Parse and validate `spec/sim_contract.yaml`. Fails loudly; never repairs."""
    path = Path(path)
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as error:
        raise SimContractError(f"{path}: cannot be read: {error}") from error

    raw = _strict_safe_load(text, path)
    if not isinstance(raw, dict):
        raise SimContractError(f"{path}: the contract must be a mapping at the top level")

    missing = [key for key in REQUIRED_SECTIONS if key not in raw]
    if missing:
        raise SimContractError(
            f"{path}: missing required section(s) {missing}. Every section is required; "
            "a missing one is an unowned semantic, not a default."
        )

    _check_closed_world(raw, "", path)

    version = _req_str(raw, "sim_contract_version", str(path))
    if version != LOCKED_SIM_CONTRACT_VERSION:
        raise SimContractError(
            f"{path}: sim_contract_version is {version!r}; this loader implements "
            f"{LOCKED_SIM_CONTRACT_VERSION!r}. A document declaring another format must not be "
            "silently consumed by a parser written for this one."
        )

    _validate_versions(raw, path)
    _validate_rng(raw, path)
    _validate_seeding(raw, path)
    _validate_components(raw, path)
    _validate_stream_assignment(raw, path)
    _validate_jump(raw, path)
    _validate_distributions(raw, path)
    _validate_cheng(raw, path)
    _validate_risk(raw, path)
    _validate_accumulation(raw, path)
    _validate_contribution(raw, path)
    _validate_kernel(raw, path)
    _validate_numerical_domain(raw, path)
    _validate_dependence(raw, path)
    _validate_publication(raw, path)
    _validate_command_surface(raw, path)
    _validate_request_fingerprint(raw, path)
    _validate_result_digest(raw, path)
    _validate_label_sets(raw, path)
    _validate_sim_state(raw, path)
    _validate_prerequisite(raw, path)
    _validate_run_id(raw, path)
    _validate_statistics(raw, path)
    _validate_contingency(raw, path)
    _validate_results_minimum(raw, path)

    layout = _parse_sim_data(raw, path)
    _validate_iterations(raw, path, layout)

    references = _parse_references(raw, path)
    _validate_reference_set(references, path)

    _forbid_seed_range(raw, path)
    _forbid_tolerance(raw, path)

    return SimContract(
        version=version,
        rng_version=int(raw["versions"]["rng_version"]),
        sim_method_version=int(raw["versions"]["sim_method_version"]),
        layout=layout,
        authority_references=references,
        raw=raw,
        source_path=path,
    )


# ---------------------------------------------------------------------------
# Section validators
# ---------------------------------------------------------------------------
def _validate_versions(raw: dict, path: Path) -> None:
    block = _map(raw, "versions", path)
    for key in ("rng_version", "sim_method_version"):
        value = _req(block, key, f"{path}: versions")
        if not isinstance(value, int) or isinstance(value, bool) or value < 1:
            raise SimContractError(
                f"{path}: versions.{key} must be a positive integer, got {value!r}"
            )
    source = _req_str(block, "result_digest_version_source", f"{path}: versions")
    if source != LOCKED_VERSION_FIELD_SOURCE:
        raise SimContractError(
            f"{path}: versions.result_digest_version_source is {source!r}; the accepted Step-0 "
            f"decision is {LOCKED_VERSION_FIELD_SOURCE!r}. The PCCM-RD version field IS "
            "SIM_METHOD_VERSION; no third version exists, and the field may not be left "
            "semantically ownerless."
        )
    bumps = _map(block, "bump_ownership", f"{path}: versions")
    _exact_sequence(
        bumps.get("rng_version"), LOCKED_RNG_VERSION_BUMPS,
        f"{path}: versions.bump_ownership.rng_version",
    )
    _exact_sequence(
        bumps.get("sim_method_version"), LOCKED_SIM_METHOD_VERSION_BUMPS,
        f"{path}: versions.bump_ownership.sim_method_version",
    )
    overlap = set(LOCKED_RNG_VERSION_BUMPS) & set(LOCKED_SIM_METHOD_VERSION_BUMPS)
    if overlap:  # pragma: no cover - guards the locked table itself
        raise SimContractError(f"{path}: a change is owned by two versions: {sorted(overlap)}")


def _validate_rng(raw: dict, path: Path) -> None:
    block = _map(raw, "rng", path)
    where = f"{path}: rng"
    family = _req_str(block, "family", where)
    if family != LOCKED_RNG_FAMILY:
        raise SimContractError(f"{where}: family must be {LOCKED_RNG_FAMILY!r}, got {family!r}")

    constants = _map(block, "constants", where)
    for key, expected in LOCKED_RNG_CONSTANTS.items():
        actual = _req(constants, key, f"{where}: constants")
        if isinstance(expected, int):
            if not isinstance(actual, int) or isinstance(actual, bool) or actual != expected:
                raise SimContractError(
                    f"{where}: constants.{key} must be exactly {expected}, got {actual!r}"
                )
        else:
            if not isinstance(actual, float) or actual != expected:
                raise SimContractError(
                    f"{where}: constants.{key} must be exactly {expected!r}, got {actual!r}. "
                    "This is a bit-exact Double, not an approximation."
                )
    extra = set(constants) - set(LOCKED_RNG_CONSTANTS)
    if extra:
        raise SimContractError(f"{where}: constants declares unknown key(s) {sorted(extra)}")

    state = _map(block, "state", where)
    if _req(state, "words", f"{where}: state") != 6:
        raise SimContractError(f"{where}: state.words must be 6")
    _exact_sequence(state.get("order"), LOCKED_STATE_ORDER, f"{where}: state.order")
    _require_value(state, "orientation", "oldest_first", f"{where}: state")
    _require_value(state, "matrix_operand_orientation", "newest_first", f"{where}: state")
    _require_true(state, "reversal_required_at_matrix_boundary", f"{where}: state")

    recurrence = _map(block, "recurrence", where)
    _require_value(recurrence, "p1", "(a12 * s11 - a13n * s10) mod m1", f"{where}: recurrence")
    _require_value(recurrence, "p2", "(a21 * s22 - a23n * s20) mod m2", f"{where}: recurrence")
    _req_str(recurrence, "advance", f"{where}: recurrence")

    combination = _map(block, "combination", where)
    _require_value(
        combination, "rule",
        "if p1 <= p2 then u = (p1 - p2 + m1) * norm else u = (p1 - p2) * norm",
        f"{where}: combination",
    )
    _require_value(
        combination, "comparison_operator", "less_than_or_equal", f"{where}: combination"
    )

    domain = _map(block, "output_domain", where)
    if (domain.get("lower"), domain.get("upper")) != (0, 1):
        raise SimContractError(f"{where}: output_domain must be lower 0, upper 1")
    _require_false(domain, "lower_inclusive", f"{where}: output_domain")
    _require_false(domain, "upper_inclusive", f"{where}: output_domain")

    arithmetic = _map(block, "arithmetic", where)
    _require_value(
        arithmetic, "modulo_semantics", "exact_non_negative_residue", f"{where}: arithmetic"
    )
    _require_value(
        arithmetic, "implementation_requirement", "exact_safe_double_path",
        f"{where}: arithmetic",
    )
    _require_false(arithmetic, "naive_floating_modulo_permitted", f"{where}: arithmetic")


def _validate_seeding(raw: dict, path: Path) -> None:
    block = _map(raw, "seeding", path)
    where = f"{path}: seeding"
    _exact_sequence(block.get("modes"), LOCKED_SEED_MODES, f"{where}: modes")
    _require_value(block, "blank_input_means", "AUTO", where)
    _require_value(block, "populated_input_means", "FIXED", where)

    scalar = _map(block, "scalar_to_state", where)
    _require_value(scalar, "rule", "repeated_scalar", f"{where}: scalar_to_state")
    _require_value(
        scalar, "expansion", "[seed, seed, seed, seed, seed, seed]",
        f"{where}: scalar_to_state",
    )
    if scalar.get("mixer") is not None:
        raise SimContractError(
            f"{where}: scalar_to_state.mixer must be null. D6-05 closed on the repeated scalar; "
            "a mixer is a new portability surface with no demonstrated requirement."
        )
    _require_false(scalar, "alternate_expansion_permitted", f"{where}: scalar_to_state")

    auto = _map(block, "auto", where)
    _require_value(auto, "modulus", LOCKED_AUTO_MODULUS, f"{where}: auto")
    _require_value(auto, "multiplier", LOCKED_AUTO_MULTIPLIER, f"{where}: auto")
    _require_value(auto, "period", LOCKED_AUTO_PERIOD, f"{where}: auto")
    _require_value(
        auto, "mapping", "effective_seed = multiplier ^ auto_nonce mod modulus", f"{where}: auto"
    )
    _require_value(auto, "mapping_kind", "modular_exponentiation", f"{where}: auto")
    _require_value(
        auto, "implementation_complexity_requirement", "O(log nonce)", f"{where}: auto"
    )
    _require_false(auto, "stepped_multiplication_is_the_authority", f"{where}: auto")
    _require_value(
        auto, "freshness_scope", "single_workbook_persisted_nonce_history", f"{where}: auto"
    )
    _require_false(auto, "cross_workbook_uniqueness_claimed", f"{where}: auto")
    _require_false(auto, "timestamp_derived_uniqueness_permitted", f"{where}: auto")

    life = _map(block, "nonce_lifecycle", where)
    lwhere = f"{where}: nonce_lifecycle"
    _require_value(life, "meaning", "next_nonce_to_allocate", lwhere)
    _require_value(life, "initial", 0, lwhere)
    _require_value(life, "first_valid_allocation", 0, lwhere)
    _require_value(life, "last_valid_allocation", LOCKED_NONCE_EXHAUSTED - 1, lwhere)
    _require_value(life, "exhausted_value", LOCKED_NONCE_EXHAUSTED, lwhere)
    _require_value(life, "on_exhaustion", "REFUSE", lwhere)
    _require_false(life, "wrap_permitted", lwhere)
    _require_false(life, "reuse_permitted", lwhere)
    _exact_sequence(life.get("order"), LOCKED_NONCE_LIFECYCLE, f"{lwhere}: order")
    _require_false(life, "failure_before_allocation_consumes_nonce", lwhere)
    _require_true(life, "failure_after_allocation_consumes_nonce", lwhere)
    _exact_sequence(
        life.get("attempt_metadata_preserves"),
        ("consumed_auto_nonce", "effective_seed"),
        f"{lwhere}: attempt_metadata_preserves",
    )
    _require_true(life, "prior_successful_publication_untouched", lwhere)


def _validate_components(raw: dict, path: Path) -> None:
    block = _map(raw, "components", path)
    where = f"{path}: components"
    kinds = _seq(block, "kinds", where)
    keys = [k.get("key") if isinstance(k, dict) else None for k in kinds]
    if len(set(keys)) != len(keys):
        raise SimContractError(f"{where}: duplicate component kind(s) in {keys}")
    _exact_sequence(keys, LOCKED_COMPONENT_KINDS, f"{where}: kinds")
    expected = {
        "COST_SAMPLE": ("cost_line", "value"),
        "RISK_OCCURRENCE": ("risk", "occurrence"),
        "RISK_SEVERITY": ("risk", "severity"),
    }
    for entry in kinds:
        key = entry["key"]
        driver_kind, role = expected[key]
        _require_value(entry, "driver_kind", driver_kind, f"{where}: {key}")
        _require_value(entry, "role", role, f"{where}: {key}")
        _require_value(entry, "per_driver", 1, f"{where}: {key}")
    _require_value(
        block, "count_rule", "cost_line_count + 2 * risk_count", where
    )


def _validate_stream_assignment(raw: dict, path: Path) -> None:
    block = _map(raw, "stream_assignment", path)
    where = f"{path}: stream_assignment"
    _require_value(block, "policy", "canonical_sorted_order", where)
    _exact_sequence(block.get("sort_keys"), LOCKED_SORT_KEYS, f"{where}: sort_keys")
    _require_value(block, "permanent_id_comparison", LOCKED_ID_COMPARISON, where)
    for flag in (
        "physical_row_order_permitted",
        "locale_order_permitted",
        "case_insensitive_order_permitted",
        "numeric_suffix_interpretation_permitted",
    ):
        _require_false(block, flag, where)
    _require_value(block, "index_origin", 0, where)
    _req_str(block, "index_rule", where)
    _require_value(
        block, "accepted_consequence",
        "inserting_or_removing_a_driver_may_change_later_stream_identities", where,
    )


def _validate_jump(raw: dict, path: Path) -> None:
    block = _map(raw, "jump", path)
    where = f"{path}: jump"
    _require_value(block, "stream_spacing_exponent", LOCKED_JUMP_EXPONENT, where)
    if block.get("substream_spacing_exponent") is not None:
        raise SimContractError(
            f"{where}: substream_spacing_exponent must be null. Phase 6 uses no substreams; "
            "an unused facility is a second contract to maintain."
        )
    _require_false(block, "substreams_used_in_phase_6", where)
    _require_value(block, "decomposition", "mult_mod_m", where)
    _require_value(block, "decomposition_h", LOCKED_JUMP_H, where)
    _require_false(block, "naive_floating_matrix_product_permitted", where)
    _check_matrix(block.get("a1_p127"), LOCKED_A1_P127, f"{where}: a1_p127")
    _check_matrix(block.get("a2_p127"), LOCKED_A2_P127, f"{where}: a2_p127")
    for key in ("a1_p127_sha256", "a2_p127_sha256"):
        digest = _req_str(block, key, where)
        if not re.fullmatch(r"[0-9a-f]{64}", digest):
            raise SimContractError(f"{where}: {key} must be 64 lowercase hex characters")


def _check_matrix(value: Any, expected: tuple, where: str) -> None:
    if not isinstance(value, list) or len(value) != 3:
        raise SimContractError(f"{where}: must be a 3 x 3 matrix, got {value!r}")
    for i, row in enumerate(value):
        if not isinstance(row, list) or len(row) != 3:
            raise SimContractError(f"{where}: row {i} must have exactly 3 elements, got {row!r}")
        for j, element in enumerate(row):
            if not isinstance(element, int) or isinstance(element, bool):
                raise SimContractError(f"{where}[{i}][{j}]: must be an integer, got {element!r}")
            if element != expected[i][j]:
                raise SimContractError(
                    f"{where}[{i}][{j}] is {element}, but the accepted Step-0 jump authority "
                    f"is {expected[i][j]}. A single altered element produces a plausible but "
                    "wrong stream."
                )


def _validate_distributions(raw: dict, path: Path) -> None:
    block = _map(raw, "distributions", path)
    where = f"{path}: distributions"
    _exact_sequence(block.get("families"), LOCKED_FAMILIES, f"{where}: families")

    degenerate = _map(block, "degenerate", where)
    dwhere = f"{where}: degenerate"
    _require_value(degenerate, "condition", "a == m == b", dwhere)
    _require_true(degenerate, "detected_before_dispatch", dwhere)
    _require_true(degenerate, "detected_before_parameterisation", dwhere)
    _require_value(degenerate, "returns", "a", dwhere)
    _require_value(degenerate, "uniforms_consumed", 0, dwhere)
    _require_false(degenerate, "sampler_entered", dwhere)
    _require_false(degenerate, "stream_state_changed", dwhere)
    _require_true(degenerate, "applies_to_all_families", dwhere)

    uniform = _map(block, "uniform", where)
    _require_false(uniform, "most_likely_used", f"{where}: uniform")
    _require_value(uniform, "transform", "x = (1 - u) * a + u * b", f"{where}: uniform")
    _require_value(uniform, "formulation", "stable_convex", f"{where}: uniform")
    _require_value(uniform, "uniforms_per_non_degenerate_sample", 1, f"{where}: uniform")

    tri = _map(block, "triangular", where)
    twhere = f"{where}: triangular"
    _require_value(tri, "method", "inverse_cdf", twhere)
    _require_value(tri, "uniforms_per_non_degenerate_sample", 1, twhere)
    _require_value(tri, "branch_point", "c = (m - a) / (b - a)", twhere)
    _require_value(
        tri, "lower_branch", "u <= c : x = a + sqrt(u * (b - a) * (m - a))", twhere
    )
    _require_value(
        tri, "upper_branch", "u >  c : x = b - sqrt((1 - u) * (b - a) * (b - m))", twhere
    )
    boundary = _map(tri, "boundary_cases", twhere)
    for key in ("m_equals_a", "m_equals_b"):
        _req_str(boundary, key, f"{twhere}: boundary_cases")
    _require_true(tri, "rng_endpoints_open", twhere)
    _require_true(tri, "normalised_formulation_required", twhere)
    _req_str(tri, "conditioning_scale", twhere)

    pert = _map(block, "beta_pert", where)
    pwhere = f"{where}: beta_pert"
    _require_value(pert, "lambda", LOCKED_PERT_LAMBDA, pwhere)
    _require_value(pert, "shape_ratio", "r = (m - a) / (b - a)", pwhere)
    _require_value(pert, "alpha", "1 + 4 * r", pwhere)
    _require_value(pert, "beta", "1 + 4 * (1 - r)", pwhere)
    _require_value(pert, "alpha_plus_beta", 6, pwhere)
    _require_value(pert, "shape_lower", 1, pwhere)
    _require_value(pert, "shape_upper", 5, pwhere)
    _require_value(pert, "rescale", "x = (1 - y) * a + y * b, y ~ Beta(alpha, beta)", pwhere)
    _require_value(pert, "rescale_formulation", "stable_convex", pwhere)
    _require_true(pert, "normalised_formulation_required", pwhere)
    dispatch = _map(pert, "dispatch", pwhere)
    _require_value(dispatch, "rule", LOCKED_DISPATCH_RULE, f"{pwhere}: dispatch")
    _require_value(
        dispatch, "comparison_operator", "strictly_greater_than", f"{pwhere}: dispatch"
    )
    _require_value(
        dispatch, "equality_belongs_to", LOCKED_DISPATCH_EQUALITY_OWNER, f"{pwhere}: dispatch"
    )


def _validate_cheng(raw: dict, path: Path) -> None:
    block = _map(raw, "cheng", path)
    where = f"{path}: cheng"
    _require_value(
        block, "uniforms_per_non_degenerate_proposal_attempt", LOCKED_UNIFORMS_PER_ATTEMPT, where
    )
    _require_value(block, "logit_form", LOCKED_LOGIT_FORM, where)
    _require_value(block, "logit_form_rejected_alternative", LOCKED_REJECTED_LOGIT_FORM, where)
    _require_true(block, "literals_are_literal", where)
    _require_false(block, "algebraic_simplification_permitted", where)

    bb = _map(block, "bb", where)
    _require_value(bb, "applies_when", "min(alpha, beta) > 1", f"{where}: bb")
    _check_orientation(bb, LOCKED_BB_ORIENTATION, f"{where}: bb")
    _exact_sequence(bb.get("per_driver"), LOCKED_BB_PER_DRIVER, f"{where}: bb.per_driver")
    _exact_sequence(bb.get("per_attempt"), LOCKED_BB_PER_ATTEMPT, f"{where}: bb.per_attempt")
    _exact_sequence(bb.get("literals"), LOCKED_BB_LITERALS, f"{where}: bb.literals")
    _require_value(bb, "acceptance_operator", "greater_than_or_equal", f"{where}: bb")
    _req_str(bb, "return", f"{where}: bb")

    bc = _map(block, "bc", where)
    _require_value(bc, "applies_when", "min(alpha, beta) <= 1", f"{where}: bc")
    _check_orientation(bc, LOCKED_BC_ORIENTATION, f"{where}: bc")
    _exact_sequence(bc.get("literals"), LOCKED_BC_LITERALS, f"{where}: bc.literals")
    _require_value(bc, "acceptance_operator", "greater_than_or_equal", f"{where}: bc")
    for key in ("per_driver", "per_attempt"):
        _seq(bc, key, f"{where}: bc")
    _req_str(bc, "return", f"{where}: bc")

    effect = _map(block, "literal_effect", where)
    _require_value(
        effect, "squeeze_literals_affect", "acceptance_decision_only", f"{where}: literal_effect"
    )
    _require_value(
        effect, "logit_form_affects", "returned_sample_value", f"{where}: literal_effect"
    )

    binding = _map(block, "source_binding", where)
    _req_str(binding, "evidence_file", f"{where}: source_binding")
    digest = _req_str(binding, "functions_sha256", f"{where}: source_binding")
    if not re.fullmatch(r"[0-9a-f]{64}", digest):
        raise SimContractError(
            f"{where}: source_binding.functions_sha256 must be 64 lowercase hex characters"
        )
    vectors = _map(block, "conformance_vectors", where)
    _req_str(vectors, "evidence_file", f"{where}: conformance_vectors")
    _require_value(vectors, "role", "conformance_authority", f"{where}: conformance_vectors")
    _require_false(vectors, "runtime_lookup_table", f"{where}: conformance_vectors")

    # The rejected alternative must not be smuggled into an expression list.
    for name, entries in (("bb", bb.get("per_attempt")), ("bc", bc.get("per_attempt"))):
        for line in entries or []:
            if "log1p" in str(line):
                raise SimContractError(
                    f"{where}: {name}.per_attempt uses log1p; the locked logit form is "
                    f"{LOCKED_LOGIT_FORM!r}, and the alternative changes the RETURNED SAMPLE, "
                    "not merely the acceptance decision."
                )


def _check_orientation(block: dict, expected: dict, where: str) -> None:
    orientation = _map(block, "orientation", where)
    for key, value in expected.items():
        actual = _req(orientation, key, f"{where}: orientation")
        if actual != value:
            raise SimContractError(
                f"{where}: orientation.{key} is {actual!r}, but the accepted formulation is "
                f"{value!r}. BB and BC orient OPPOSITELY; inverting one is a silent defect "
                "that returns a valid Beta variate of the mirrored distribution."
            )


def _validate_risk(raw: dict, path: Path) -> None:
    block = _map(raw, "risk", path)
    where = f"{path}: risk"
    occurrence = _map(block, "occurrence", where)
    owhere = f"{where}: occurrence"
    _require_value(occurrence, "uniforms_per_risk_per_iteration", 1, owhere)
    _require_value(occurrence, "rule", "occurred = u_occurrence < probability", owhere)
    _require_value(occurrence, "comparison_operator", "strictly_less_than", owhere)
    _require_true(occurrence, "probability_zero_never_occurs", owhere)
    _require_true(occurrence, "probability_one_always_occurs", owhere)

    severity = _map(block, "severity", where)
    swhere = f"{where}: severity"
    policy = _req_str(severity, "invocation_policy", swhere)
    if policy != "unconditional":
        raise SimContractError(
            f"{swhere}: invocation_policy is {policy!r}; D6-18 closed on 'unconditional'. "
            "A conditional policy destroys probability-only scenario comparability."
        )
    _require_true(severity, "sampler_invoked_every_risk_iteration", swhere)
    _require_value(
        severity, "non_degenerate_consumption",
        "as_the_selected_sampler_contract_requires", swhere,
    )
    _require_value(severity, "degenerate_consumption", 0, swhere)
    _require_false(severity, "degenerate_stream_state_changed", swhere)
    _require_true(severity, "value_used_only_when_occurred", swhere)

    _require_false(block, "probability_folded_into_knom", where)
    _require_false(block, "probability_folded_into_kpv", where)

    # The withdrawn phrase is false for a degenerate severity and must not appear
    # anywhere in the contract, in any casing or spacing.
    flat = _flatten_text(raw).lower()
    for phrase in ("stream advances once per iteration", "advances once per iteration"):
        if phrase in flat:
            raise SimContractError(
                f"{path}: the contract contains {phrase!r}. That wording was WITHDRAWN in "
                "Step-0 section 6.1: it is false for a degenerate severity, which is invoked "
                "every iteration and advances nothing."
            )


def _validate_accumulation(raw: dict, path: Path) -> None:
    block = _map(raw, "accumulation", path)
    where = f"{path}: accumulation"
    _exact_sequence(
        block.get("driver_kind_order"), ("cost_line", "risk"), f"{where}: driver_kind_order"
    )
    _require_value(block, "within_kind_order", "ascending_permanent_id", where)
    _require_value(block, "permanent_id_comparison", LOCKED_ID_COMPARISON, where)
    _require_false(block, "physical_row_order_permitted", where)
    _exact_sequence(block.get("accumulators"), ("nominal", "pv"), f"{where}: accumulators")
    _require_true(block, "accumulators_share_driver_order", where)


def _validate_contribution(raw: dict, path: Path) -> None:
    """What a driver actually contributes to one iteration.

    Without this, an engine could satisfy every other section and still sample
    total cost instead of unit cost, drop Quantity, apply it twice, or discount
    the nominal total instead of computing PV independently.
    """
    block = _map(raw, "contribution", path)
    where = f"{path}: contribution"

    cost = _map(block, "cost_line", where)
    cwhere = f"{where}: cost_line"
    _require_value(cost, "sampled_quantity", "unit_cost", cwhere)
    _require_value(cost, "sampled_from", "distribution(Min, MostLikely, Max)", cwhere)
    _require_false(cost, "total_cost_uncertainty_sampled", cwhere)
    _require_true(cost, "quantity_is_deterministic", cwhere)
    _require_false(cost, "quantity_inside_distribution", cwhere)
    _require_value(cost, "quantity_applications", 1, cwhere)
    _require_false(cost, "probability_applies", cwhere)
    _require_value(cost, "nominal", "unit_cost * Quantity * Knom", cwhere)
    _require_value(cost, "pv", "unit_cost * Quantity * Kpv", cwhere)

    risk = _map(block, "risk", where)
    rwhere = f"{where}: risk"
    _require_value(risk, "occurred", "occurrence_uniform < Probability", rwhere)
    _require_value(risk, "severity_source", "severity_sampler_under_d6_18b", rwhere)
    _require_false(risk, "quantity_applies", rwhere)
    _require_false(risk, "probability_folded_into_k_factors", rwhere)
    _require_false(risk, "occurrence_and_severity_share_a_stream", rwhere)
    _require_value(risk, "nominal_when_occurred", "severity * Knom", rwhere)
    _require_value(risk, "pv_when_occurred", "severity * Kpv", rwhere)
    _require_value(risk, "nominal_when_not_occurred", 0, rwhere)
    _require_value(risk, "pv_when_not_occurred", 0, rwhere)

    _require_false(block, "pv_derived_from_nominal", where)

    total = _map(block, "iteration_total", where)
    _require_value(
        total, "rule",
        "canonical sum of every Cost Line contribution, then every Risk contribution",
        f"{where}: iteration_total",
    )
    _require_value(total, "order_source", "accumulation", f"{where}: iteration_total")
    _require_true(total, "measures_independent", f"{where}: iteration_total")


def _validate_kernel(raw: dict, path: Path) -> None:
    """The inherited hot-kernel boundary. Semantic authority, not a note."""
    block = _map(raw, "kernel", path)
    where = f"{path}: kernel"
    _require_true(block, "inputs_resolved_once_before_simulation", where)
    _require_true(block, "operates_on_resolved_in_memory_structures", where)
    for flag in (
        "worksheet_access_inside_iteration_loop",
        "range_access_inside_iteration_loop",
        "listobject_access_inside_iteration_loop",
        "application_object_access_inside_iteration_loop",
        "thisworkbook_or_activeworkbook_access_inside_iteration_loop",
        "com_round_trip_inside_iteration_loop",
        "recomputes_worksheet_inflation_inside_loop",
        "recomputes_worksheet_fx_inside_loop",
        "recomputes_worksheet_profiles_inside_loop",
    ):
        _require_false(block, flag, where)
    _exact_sequence(
        block.get("resolved_before_loop"),
        ("knom_per_driver", "kpv_per_driver", "quantities", "probabilities",
         "distribution_parameters", "component_stream_initial_states"),
        f"{where}: resolved_before_loop",
    )


def _validate_numerical_domain(raw: dict, path: Path) -> None:
    """Plan section 4.6 in full. Phase 6 may not silently narrow Phase 5's domain."""
    block = _map(raw, "numerical_domain", path)
    where = f"{path}: numerical_domain"
    _require_true(block, "negative_values_legal", where)
    _require_true(block, "supports_crossing_zero_legal", where)
    for key in ("positivity_rule", "magnitude_restriction"):
        if block.get(key) is not None:
            raise SimContractError(
                f"{where}: {key} must be null. Phase 5 accepted any finite, correctly ordered "
                "triple and built overflow-safe primitives rather than restricting the domain; "
                "Phase 6 inherits that domain and may not narrow it."
            )
    _require_false(block, "narrower_than_phase5", where)
    _require_false(
        block, "representable_result_refused_for_naive_intermediate_overflow", where
    )
    _require_value(
        block, "refusal_when_no_valid_double_result",
        "explicit_and_names_the_numerical_stage", where,
    )
    _require_false(block, "silent_non_finite_result_permitted", where)
    disciplines = _map(block, "disciplines", where)
    for key, value in (
        ("driver_contribution", "accepted_safe_product"),
        ("accumulation", "accepted_safe_signed_sum"),
        ("percentile_interpolation", "convex"),
        ("contingency_subtraction", "accepted_safe_subtract"),
        ("statistics", "scale_safe"),
    ):
        _require_value(disciplines, key, value, f"{where}: disciplines")


def _validate_dependence(raw: dict, path: Path) -> None:
    """The component-stream architecture makes independence achievable. This
    STATES it, because an unstated invariant is one a later phase can break."""
    block = _map(raw, "dependence", path)
    where = f"{path}: dependence"
    _require_value(block, "inter_driver_dependence", "independent", where)
    _require_false(block, "correlation_matrix_supported", where)
    _require_false(block, "copula_supported", where)
    _require_false(block, "shared_or_hidden_dependence_permitted", where)
    _req_str(block, "authority", where)


def _validate_publication(raw: dict, path: Path) -> None:
    block = _map(raw, "publication", path)
    where = f"{path}: publication"
    _require_value(block, "persisted_source_of_truth", "_SimData", where)
    _require_value(block, "results_derives_from", "_SimData", where)
    _require_false(block, "results_recomputes_monte_carlo", where)
    _require_true(block, "publish_only_after_simulation_and_statistics_complete", where)
    _require_true(block, "commit_last", where)
    _require_false(block, "partial_new_distribution_published_on_refusal_or_failure", where)
    _require_true(block, "prior_successful_publication_survives", where)


def _validate_command_surface(raw: dict, path: Path) -> None:
    block = _map(raw, "command_surface", path)
    where = f"{path}: command_surface"
    _require_value(block, "automation_endpoint", "PCCM_RunSimulation", where)
    for flag in (
        "user_facing_run_button_in_phase_6",
        "msgbox_introduced_by_phase_6",
        "userform_introduced_by_phase_6",
        "ribbon_introduced_by_phase_6",
        "read_accessor_names_settled",
    ):
        _require_false(block, flag, where)
    interruption = _map(raw, "interruption", path)
    _require_false(
        interruption, "user_cancellation_supported_in_phase_6", f"{path}: interruption"
    )


def _validate_request_fingerprint(raw: dict, path: Path) -> None:
    block = _map(raw, "request_fingerprint", path)
    where = f"{path}: request_fingerprint"
    _exact_sequence(
        block.get("section_order"), LOCKED_SECTION_ORDER, f"{where}: section_order"
    )
    _exact_sequence(
        block.get("analytical_prefix"), LOCKED_ANALYTICAL_PREFIX, f"{where}: analytical_prefix"
    )
    order = tuple(block.get("section_order") or ())
    if order[: len(LOCKED_ANALYTICAL_PREFIX)] != LOCKED_ANALYTICAL_PREFIX:
        raise SimContractError(
            f"{where}: the accepted Phase-5 sections must remain a PREFIX of the request "
            "stream. Reordering them changes existing bytes and breaks every stored "
            "analytical fingerprint."
        )
    _require_value(block, "extension_semantics", "prefix_plus_extension", where)
    _require_false(block, "existing_sections_modified", where)

    sim = _map(block, "sim_section", where)
    _require_value(sim, "name", "SIM", f"{where}: sim_section")
    _exact_sequence(sim.get("fields"), LOCKED_SIM_FIELDS, f"{where}: sim_section.fields")
    _require_value(sim, "supplied_seed_present_only_when", "FIXED", f"{where}: sim_section")
    excluded = tuple(sim.get("excluded_fields") or ())
    missing = [f for f in LOCKED_SIM_EXCLUDED if f not in excluded]
    if missing:
        raise SimContractError(
            f"{where}: sim_section.excluded_fields omits {missing}. Each of those changes on a "
            "legitimate AUTO re-run of the same question, so including one would make every "
            "such run STALE."
        )
    overlap = set(sim.get("fields") or ()) & set(LOCKED_SIM_EXCLUDED)
    if overlap:
        raise SimContractError(
            f"{where}: sim_section.fields includes excluded field(s) {sorted(overlap)}"
        )
    _require_false(sim, "analytical_fingerprint_hashed_as_a_field", f"{where}: sim_section")
    _require_true(block, "auto_blank_seed_remains_recomputable", where)


def _validate_result_digest(raw: dict, path: Path) -> None:
    block = _map(raw, "result_digest", path)
    where = f"{path}: result_digest"
    _require_value(block, "stream_tag", LOCKED_DIGEST_TAG, where)
    _require_value(block, "version_field_source", LOCKED_VERSION_FIELD_SOURCE, where)
    _require_value(block, "section_name", LOCKED_DIGEST_SECTION, where)
    _require_value(block, "record_field_count", len(LOCKED_DIGEST_RECORD_FIELDS), where)
    _exact_sequence(
        block.get("record_fields"), LOCKED_DIGEST_RECORD_FIELDS, f"{where}: record_fields"
    )
    _exact_sequence(
        block.get("field_types"), LOCKED_DIGEST_FIELD_TYPES, f"{where}: field_types"
    )
    _require_value(block, "iteration_index_origin", 1, where)
    grammar = _map(block, "grammar", where)
    for key in ("stream", "section", "record"):
        _req_str(grammar, key, f"{where}: grammar")
    _require_value(block, "order_source", "persisted_iteration_order", where)
    _require_false(block, "samples_sorted_for_digest", where)
    _require_value(block, "equality", "exact", where)
    if block.get("tolerance") is not None:
        raise SimContractError(
            f"{where}: tolerance must be null. Digest equality is exact; an approximate "
            "comparison here would defeat the identity the digest exists to provide."
        )


def _validate_label_sets(raw: dict, path: Path) -> None:
    block = _map(raw, "label_sets", path)
    where = f"{path}: label_sets"
    _exact_sequence(block.get("sim_state"), LOCKED_SIM_STATES, f"{where}: sim_state")
    _exact_sequence(
        block.get("attempt_result"), LOCKED_ATTEMPT_RESULTS, f"{where}: attempt_result"
    )
    _exact_sequence(block.get("seed_mode"), LOCKED_SEED_MODES, f"{where}: seed_mode")


def _validate_sim_state(raw: dict, path: Path) -> None:
    """The corrected derivation: ordered, total, and blind to attempt history."""
    block = _map(raw, "sim_state", path)
    where = f"{path}: sim_state"
    states = tuple(block.get("states") or ())
    if states != LOCKED_SIM_STATES:
        raise SimContractError(
            f"{where}: states must be exactly {list(LOCKED_SIM_STATES)}, got {list(states)}. "
            "There is no fourth simulation state; revision 1's UNSELECTED was rejected."
        )

    derivation = _map(block, "derivation", where)
    dwhere = f"{where}: derivation"
    _require_true(derivation, "ordered", dwhere)
    rules = _seq(derivation, "rules", dwhere)
    actual = tuple(
        (rule.get("order"), rule.get("condition"), rule.get("status"))
        if isinstance(rule, dict) else rule
        for rule in rules
    )
    if actual != LOCKED_SIM_STATE_RULES:
        raise SimContractError(
            f"{dwhere}: the derivation must be exactly {[list(r) for r in LOCKED_SIM_STATE_RULES]}, "
            f"in that order, got {[list(r) if isinstance(r, tuple) else r for r in actual]}. "
            "The order is what makes the rules mutually exclusive, and rules 1 and 2 are what "
            "make them total: without them a corrected-then-restored request has no state at all."
        )

    # The whole point of the correction.
    _require_false(block, "attempt_result_participates_in_derivation", where)
    _require_true(block, "attempt_axis_is_orthogonal", where)
    for rule in rules:
        condition = str(rule.get("condition", ""))
        # The prohibited concepts are the ATTEMPT-RESULT labels. "successful
        # snapshot" is a different thing - the stored result the fingerprint is
        # compared against - so `success` is deliberately not in this list.
        for token in ("attempt", "refused", "failed"):
            if token in condition.lower():
                raise SimContractError(
                    f"{dwhere}: rule condition {condition!r} reads the attempt history. The "
                    "attempt result is an ORTHOGONAL audit axis and must not decide the "
                    "derived status - that is the authority conflict this correction fixes."
                )

    if block.get("no_success_valid_status") is not None:
        raise SimContractError(
            f"{where}: no_success_valid_status must be null. A blank status is the ABSENCE of a "
            "successful-comparison state, not a fourth label; naming it would create one."
        )
    _require_true(
        block, "status_evaluated_at_may_be_populated_while_status_is_blank", where
    )

    definitions = _map(block, "definitions", where)
    if tuple(definitions) != LOCKED_SIM_STATES:
        raise SimContractError(
            f"{where}: definitions must define exactly the three states, in order"
        )
    for name, text in definitions.items():
        for token in ("attempt", "refused", "failed"):
            if token in str(text).lower():
                raise SimContractError(
                    f"{where}: the definition of {name} mentions the attempt history"
                )

    failure = _map(block, "on_failure", where)
    fwhere = f"{where}: on_failure"
    _require_true(failure, "prior_sim_data_preserved", fwhere)
    _require_true(failure, "prior_results_publication_preserved", fwhere)
    _require_true(failure, "attempt_metadata_updated", fwhere)
    _require_false(failure, "partial_distribution_published", fwhere)


def derive_sim_status(
    prerequisites_resolve: bool,
    successful_snapshot_exists: bool,
    request_fingerprint_matches: bool,
) -> str | None:
    """The contract's derivation, as a total function. Returns None for BLANK.

    This is SEMANTICS, not simulation: it advances no state, draws nothing and
    reads no workbook. It exists so the truth table can be tested as a function
    rather than re-derived by prose in every test.
    """
    if not prerequisites_resolve:
        return "INVALID"
    if not successful_snapshot_exists:
        return None
    return "CURRENT" if request_fingerprint_matches else "STALE"


def _validate_prerequisite(raw: dict, path: Path) -> None:
    block = _map(raw, "prerequisite", path)
    where = f"{path}: prerequisite"
    _require_value(block, "phase5_analytical_state_required", "CURRENT", where)
    _require_false(block, "silent_recalculation_permitted", where)
    _require_false(block, "phase6_may_call_pccm_calculate", where)


def _validate_run_id(raw: dict, path: Path) -> None:
    block = _map(raw, "run_id", path)
    where = f"{path}: run_id"
    _require_value(block, "initial", 0, where)
    _require_value(block, "first_successful_value", 1, where)
    _require_value(block, "allocated_on", "successful_commit_only", where)
    _require_false(block, "failure_consumes", where)
    _require_value(block, "maximum", LOCKED_RUN_ID_MAX, where)
    _require_value(block, "on_exhaustion", "REFUSE_BEFORE_COMMIT", where)
    _require_false(block, "wrap_permitted", where)
    _require_false(block, "reuse_permitted", where)
    _require_true(block, "persisted", where)
    _exact_sequence(
        block.get("independent_of"),
        ("auto_nonce", "effective_seed", "request_fingerprint", "result_digest"),
        f"{where}: independent_of",
    )


def _validate_statistics(raw: dict, path: Path) -> None:
    block = _map(raw, "statistics", path)
    where = f"{path}: statistics"
    mean = _map(block, "mean", where)
    _require_value(mean, "method", "sample_mean", f"{where}: mean")
    _require_true(mean, "scale_safe_required", f"{where}: mean")

    sd = _map(block, "standard_deviation", where)
    _require_value(sd, "method", "sample_standard_deviation", f"{where}: standard_deviation")
    _require_value(sd, "divisor", LOCKED_SD_DIVISOR, f"{where}: standard_deviation")
    _require_false(sd, "naive_sum_of_squares_permitted", f"{where}: standard_deviation")

    pct = _map(block, "percentile", where)
    _require_value(pct, "method", LOCKED_PERCENTILE_METHOD, f"{where}: percentile")
    formula = _map(pct, "formula", f"{where}: percentile")
    for key, value in (
        ("h", "(n - 1) * p"),
        ("lo", "floor(h)"),
        ("hi", "min(lo + 1, n - 1)"),
        ("f", "h - lo"),
        ("value", "(1 - f) * x[lo] + f * x[hi]"),
    ):
        _require_value(formula, key, value, f"{where}: percentile.formula")
    _require_value(pct, "interpolation", "convex", f"{where}: percentile")

    _require_value(block, "sorting", "on_copies_only", where)
    _exact_sequence(block.get("measures"), ("nominal", "pv"), f"{where}: measures")

    # The ladder is retained BY REFERENCE. Its values are not copied here, so a
    # legitimate future change to the owner's list flows through one authority
    # instead of requiring a duplicate to be edited in step.
    _exact_sequence(
        block.get("fixed_nonselectable_percentiles"), ("P10",),
        f"{where}: fixed_nonselectable_percentiles",
    )
    _require_true(block, "include_all_selectable_ladder_values", where)
    _require_value(block, "selectable_ladder_owner", "input_contract.yaml", where)
    _require_value(
        block, "selectable_ladder_locator", "config_tables.confidence_levels", where
    )
    _exact_sequence(
        block.get("headline_percentiles"), ("P10", "P50", "P70", "P90"),
        f"{where}: headline_percentiles",
    )
    _exact_sequence(
        block.get("moments_and_extremes"),
        ("mean", "sample_standard_deviation", "minimum", "maximum"),
        f"{where}: moments_and_extremes",
    )
    _require_false(block, "p10_selectable", where)
    scl = _map(block, "selected_confidence_level", where)
    _require_value(scl, "role", "reporting_selector", f"{where}: selected_confidence_level")
    for flag in ("enters_simulation_execution", "enters_request_fingerprint",
                 "affects_staleness"):
        _require_false(scl, flag, f"{where}: selected_confidence_level")


def _validate_contingency(raw: dict, path: Path) -> None:
    block = _map(raw, "contingency", path)
    where = f"{path}: contingency"
    _exact_sequence(block.get("measures"), ("nominal", "pv"), f"{where}: measures")
    _require_value(
        block, "formula", "selected_px_total - deterministic_base_estimate_a", where
    )
    _require_value(block, "baseline", LOCKED_CONTINGENCY_BASELINE, where)
    _exact_sequence(
        block.get("forbidden_baselines"), LOCKED_FORBIDDEN_BASELINES,
        f"{where}: forbidden_baselines",
    )
    _require_false(block, "workbook_recommends_a_confidence_level", where)


def _validate_results_minimum(raw: dict, path: Path) -> None:
    block = _map(raw, "results_minimum", path)
    where = f"{path}: results_minimum"
    _exact_sequence(
        block.get("sections"), ("Run Stamp", "Summary Statistics"), f"{where}: sections"
    )
    deferred = tuple(block.get("deferred") or ())
    for required in ("Annual Cash Flow", "Reconciliation presentation", "Dashboard",
                     "Charts", "Sensitivity"):
        if required not in deferred:
            raise SimContractError(f"{where}: deferred omits {required!r}")
    _require_false(block, "annual_simulated_samples_contracted", where)


# ---------------------------------------------------------------------------
# _SimData layout and the derived ceiling - D6-08
# ---------------------------------------------------------------------------
def _parse_sim_data(raw: dict, path: Path) -> SimDataLayout:
    block = _map(raw, "sim_data", path)
    where = f"{path}: sim_data"
    sheet = _req_str(block, "sheet", where)
    if sheet != LOCKED_SIM_DATA_SHEET:
        raise SimContractError(f"{where}: sheet must be {LOCKED_SIM_DATA_SHEET!r}")
    visibility = _req_str(block, "required_visibility", where)
    if visibility != "veryHidden":
        raise SimContractError(
            f"{where}: required_visibility must be 'veryHidden'; _SimData is machine data "
            "with no audit value in raw form."
        )

    reserved_raw = _seq(block, "reserved_rows", where)
    reserved: list[tuple[int, int, str]] = []
    for i, entry in enumerate(reserved_raw):
        if not isinstance(entry, dict):
            raise SimContractError(f"{where}: reserved_rows[{i}] must be a mapping")
        rows = entry.get("rows")
        if (not isinstance(rows, list) or len(rows) != 2
                or not all(isinstance(v, int) and not isinstance(v, bool) for v in rows)):
            raise SimContractError(
                f"{where}: reserved_rows[{i}].rows must be a two-element [first, last] "
                f"integer range, got {rows!r}"
            )
        first, last = rows
        if first < 1 or last < first:
            raise SimContractError(f"{where}: reserved_rows[{i}] range {rows!r} is not ordered")
        reserved.append((first, last, _req_str(entry, "purpose", f"{where}: reserved_rows[{i}]")))

    # Contiguous from row 1, no gap and no overlap. A gap would be an
    # unaccounted-for row and an overlap would double-count one, and either makes
    # H unauditable.
    expected_next = 1
    for first, last, purpose in reserved:
        if first != expected_next:
            raise SimContractError(
                f"{where}: reserved_rows must tile rows 1..H with no gap and no overlap; "
                f"expected the next range to start at row {expected_next}, but "
                f"{purpose!r} starts at {first}"
            )
        expected_next = last + 1
    h = expected_next - 1
    if h < 1:
        raise SimContractError(f"{where}: reserved_rows must cover at least row 1")

    records = _map(block, "iteration_records", where)
    rwhere = f"{where}: iteration_records"
    header_row = _req_int(records, "header_row", rwhere)
    first_row = _req_int(records, "first_iteration_row", rwhere)
    footer = _req(records, "footer_rows", rwhere)
    if footer != 0:
        raise SimContractError(
            f"{rwhere}: footer_rows must be 0. A footer silently reduces capacity below the "
            "declared ceiling, and the ceiling is what a pre-flight refusal is computed from."
        )
    if header_row != h:
        raise SimContractError(
            f"{rwhere}: header_row is {header_row}, but reserved_rows account for {h} rows. "
            "The header is the last reserved row; any other value means a row is either "
            "unaccounted for or counted twice."
        )
    if first_row != header_row + 1:
        raise SimContractError(
            f"{rwhere}: first_iteration_row must be header_row + 1 = {header_row + 1}, "
            f"got {first_row}"
        )
    _require_value(records, "order", "canonical_iteration_order", rwhere)
    _require_false(records, "sorted", rwhere)

    columns = _seq(records, "columns", rwhere)
    keys = tuple(c.get("key") if isinstance(c, dict) else None for c in columns)
    _exact_sequence(keys, LOCKED_ITERATION_COLUMNS, f"{rwhere}: columns")
    seen_columns: set[str] = set()
    for entry in columns:
        column = _req_str(entry, "column", rwhere)
        if not re.fullmatch(r"[A-Z]{1,3}", column):
            raise SimContractError(f"{rwhere}: column {column!r} is not a column letter")
        if column in seen_columns:
            raise SimContractError(f"{rwhere}: column {column!r} is declared twice")
        seen_columns.add(column)
        _req_str(entry, "header", rwhere)
        _req_str(entry, "value_type", rwhere)

    identity = _map(block, "run_identity", where)
    iwhere = f"{where}: run_identity"
    first_field_row = _req_int(identity, "first_row", iwhere)
    last_field_row = _req_int(identity, "last_row", iwhere)
    fields = _seq(identity, "fields", iwhere)
    rows_declared = [_req_int(f, "row", iwhere) for f in fields]
    if rows_declared != list(range(first_field_row, last_field_row + 1)):
        raise SimContractError(
            f"{iwhere}: field rows must be exactly {first_field_row}..{last_field_row}, "
            f"contiguous and in order, got {rows_declared}"
        )
    if last_field_row >= header_row:
        raise SimContractError(
            f"{iwhere}: the run-identity block ends at row {last_field_row}, at or below the "
            f"iteration header row {header_row}"
        )
    # EXACT, not "contains the required fields". An invented field appended on
    # the next free row was previously accepted, and a snapshot layout that can
    # grow by accident is one whose meaning drifts silently.
    actual = tuple(
        (
            _req_str(f, "key", iwhere),
            _req_int(f, "row", iwhere),
            _req_str(f, "group", iwhere),
            _req_str(f, "value_type", iwhere),
        )
        for f in fields
    )
    if actual != LOCKED_RUN_IDENTITY:
        extra = [a[0] for a in actual if a[0] not in {r[0] for r in LOCKED_RUN_IDENTITY}]
        missing = [r[0] for r in LOCKED_RUN_IDENTITY if r[0] not in {a[0] for a in actual}]
        raise SimContractError(
            f"{iwhere}: the run-identity block must be exactly the accepted layout - key, row, "
            f"group and value type, in order. unexpected={extra} missing={missing}. "
            "The persisted simulation identity is exact authority, not an extensible list."
        )
    field_keys = [a[0] for a in actual]
    label_sets = raw.get("label_sets") or {}
    for entry in fields:
        if entry.get("value_type") == "enum":
            name = _req_str(entry, "enum", iwhere)
            if name not in label_sets:
                raise SimContractError(
                    f"{iwhere}: field {entry.get('key')!r} names enum {name!r}, which is not "
                    "declared in label_sets"
                )

    excluded = tuple(block.get("excluded") or ())
    for required in LOCKED_SIM_DATA_EXCLUDED:
        if required not in excluded:
            raise SimContractError(
                f"{where}: excluded omits {required!r}. The accepted plan defers that "
                "retention; leaving it undeclared invites it back by accident."
            )

    return SimDataLayout(
        sheet=sheet,
        required_visibility=visibility,
        reserved_rows=tuple(reserved),
        header_row=header_row,
        first_iteration_row=first_row,
        footer_rows=0,
        reserved_row_count=h,
    )


def _validate_iterations(raw: dict, path: Path, layout: SimDataLayout) -> None:
    block = _map(raw, "iterations", path)
    where = f"{path}: iterations"
    _require_value(block, "business_minimum_owner", "input_contract.yaml", where)
    if "business_minimum" in block:
        raise SimContractError(
            f"{where}: business_minimum must not appear here. The >= 1000 rule is owned by "
            "input_contract.yaml; a second copy is the drift this boundary prevents."
        )
    if block.get("business_maximum") is not None:
        raise SimContractError(
            f"{where}: business_maximum must be null. No business maximum exists, and the "
            "technical storage ceiling must never be presented as one."
        )

    ceiling = _map(block, "technical_ceiling", where)
    cwhere = f"{where}: technical_ceiling"
    _require_value(ceiling, "max_excel_rows", MAX_EXCEL_ROWS, cwhere)
    declared_h = _req_int(ceiling, "reserved_rows_h", cwhere)
    declared_max = _req_int(ceiling, "max_iterations_representable", cwhere)
    if declared_h != layout.reserved_row_count:
        raise SimContractError(
            f"{cwhere}: reserved_rows_h is {declared_h}, but the declared _SimData layout "
            f"reserves {layout.reserved_row_count} rows. H is DERIVED from the layout; it is "
            "not a free constant."
        )
    if declared_max != layout.max_iterations_representable:
        raise SimContractError(
            f"{cwhere}: max_iterations_representable is {declared_max}, but "
            f"{MAX_EXCEL_ROWS} - {layout.reserved_row_count} = "
            f"{layout.max_iterations_representable}. D6-08 is closed by derivation from the "
            "layout, not by a literal that happens to look plausible."
        )
    _require_value(ceiling, "refusal_kind", "technical", cwhere)
    _require_false(ceiling, "presented_as_business_validation", cwhere)
    _exact_sequence(
        ceiling.get("refusal_precedes"),
        ("sample_allocation", "stream_construction", "auto_seed_allocation", "any_random_draw"),
        f"{cwhere}: refusal_precedes",
    )
    _require_false(ceiling, "consumes_auto_nonce", cwhere)


# ---------------------------------------------------------------------------
# Exclusions
# ---------------------------------------------------------------------------
def _forbid_seed_range(raw: dict, path: Path) -> None:
    """The admissible seed RANGE is owned by input_contract.yaml, not by this file.

    `2147483646` legitimately appears twice here as the AUTO cycle period and the
    nonce exhaustion point, which are facts about the NONCE cycle and not about
    the admissible input domain. Those two locations are allowed by path; every
    other occurrence, and every seed-range-shaped key anywhere, is refused.
    """
    for location, value in _walk(raw):
        text = str(value)
        if location in SEED_RANGE_ALLOWED_LOCATIONS:
            continue
        key = location[-1] if location else ""
        lowered = str(key).lower()
        for token in SEED_RANGE_TOKENS:
            if token == "2147483646":
                if text == token:
                    raise SimContractError(
                        f"{path}: {'.'.join(str(p) for p in location)} carries the seed-domain "
                        "maximum. The admissible Random Seed domain is owned by "
                        "input_contract.yaml (D6-19a) and referenced here, never copied."
                    )
            elif token in lowered:
                raise SimContractError(
                    f"{path}: {'.'.join(str(p) for p in location)} names a seed range. That "
                    "rule belongs to input_contract.yaml and must not be duplicated here."
                )


def _forbid_tolerance(raw: dict, path: Path) -> None:
    """No comparison tolerance may live in this contract - not even a null one.

    Step-0 section 10 settled the tolerance as an evidence policy with a single
    owner OUTSIDE the contract. A `tolerance: null` field was previously allowed
    by a special case; that has been removed. Null is not an approximate value,
    but the FIELD is still a tolerance semantic sitting in the wrong file, and
    the declared boundary says there is none here. `result_digest.equality:
    exact` already states the runtime rule.

    Comments may explain where the tolerance lives. The parsed contract may not
    contain the semantic.
    """
    for location, value in _walk(raw):
        key = str(location[-1] if location else "").lower()
        for token in TOLERANCE_TOKENS:
            if token in key:
                raise SimContractError(
                    f"{path}: {'.'.join(str(p) for p in location)} declares a comparison "
                    "tolerance. Step 0 settled the tolerance as an ORACLE/EVIDENCE policy owned "
                    "outside this contract; the engine performs no approximate comparison at "
                    "run time, and a null-valued field is still the semantic in the wrong file."
                )
        if isinstance(value, str) and "tolerance" in value.lower():
            raise SimContractError(
                f"{path}: {'.'.join(str(p) for p in location)} mentions a tolerance value. "
                "See Step-0 section 10: the tolerance is not a runtime contract."
            )


# ---------------------------------------------------------------------------
# Authority references
# ---------------------------------------------------------------------------
def _parse_references(raw: dict, path: Path) -> tuple[AuthorityReference, ...]:
    entries = _seq(raw, "authority_references", str(path))
    out = []
    for i, entry in enumerate(entries):
        where = f"{path}: authority_references[{i}]"
        if not isinstance(entry, dict):
            raise SimContractError(f"{where}: must be a mapping")
        extra = set(entry) - ALLOWED_AUTHORITY_REFERENCE_KEYS
        if extra:
            raise SimContractError(f"{where}: unknown key(s) {sorted(extra)}")
        out.append(
            AuthorityReference(
                concept=_req_str(entry, "concept", where),
                owner=_req_str(entry, "owner", where),
                locator=_req_str(entry, "locator", where),
            )
        )
    return tuple(out)


def _validate_reference_set(references: tuple, path: Path) -> None:
    declared = tuple((r.concept, r.owner, r.locator) for r in references)
    if len(set(declared)) != len(declared):
        raise SimContractError(f"{path}: duplicate authority reference(s) declared")
    concepts = [r.concept for r in references]
    if len(set(concepts)) != len(concepts):
        raise SimContractError(
            f"{path}: a concept is declared twice with different owners or locators. "
            "A concept has exactly one owning authority."
        )
    if set(declared) != set(REQUIRED_AUTHORITY_REFERENCES):
        missing = sorted(set(REQUIRED_AUTHORITY_REFERENCES) - set(declared))
        extra = sorted(set(declared) - set(REQUIRED_AUTHORITY_REFERENCES))
        raise SimContractError(
            f"{path}: the authority-reference set does not match the accepted Step-1 boundary "
            f"set. missing={missing} unexpected={extra}. A deleted reference silently drops "
            "that boundary; a renamed concept or a moved locator silently redirects it."
        )


def validate_sim_against(
    sim: SimContract,
    spec: WorkbookSpec,
    contract: InputContract,
    drivers: DriverContract,
    structure: StructureContract,
    calc_document: dict[str, Any] | None = None,
) -> None:
    """Resolve every declared boundary against the authority that owns it.

    Nothing here copies a borrowed value. The reference is checked to RESOLVE, so
    a renamed or deleted upstream key fails the build instead of leaving a stale
    pointer behind.
    """
    owners: dict[str, Any] = {
        "workbook.yaml": _document_of(spec),
        "input_contract.yaml": _document_of(contract),
        "driver_contract.yaml": _document_of(drivers),
        "structure_contract.yaml": _document_of(structure),
    }
    if calc_document is not None:
        owners["calc_contract.yaml"] = calc_document

    for reference in sim.authority_references:
        if reference.owner not in owners:
            if reference.owner == "calc_contract.yaml":
                continue  # not supplied by this caller; the load-time set check still applies
            raise SimContractError(
                f"{sim.source_path}: authority reference {reference.concept!r} names owner "
                f"{reference.owner!r}, which is not one of the accepted specifications"
            )
        document = owners[reference.owner]
        node: Any = document
        for part in reference.locator.split("."):
            node = _step(node, part)
            if node is _MISSING:
                raise SimContractError(
                    f"{sim.source_path}: authority reference {reference.concept!r} points at "
                    f"{reference.owner}:{reference.locator}, which does not resolve (failed at "
                    f"{part!r}). References borrow values; they never copy them."
                )

    # ---- CONTENT bindings ------------------------------------------------
    # A locator that merely reaches SOME node is not enough where this contract
    # also restates or depends on that node's content: `_SimData` visibility could
    # be changed from veryHidden to hidden, and the distribution master list could
    # be changed outright, and both were accepted. Those were false bindings.
    workbook_doc = owners["workbook.yaml"]

    sheets = workbook_doc.get("sheets") if isinstance(workbook_doc, dict) else None
    sim_sheet = _step(sheets, sim.layout.sheet)
    if sim_sheet is _MISSING or not isinstance(sim_sheet, dict):
        raise SimContractError(
            f"{sim.source_path}: workbook.yaml declares no sheet named {sim.layout.sheet!r}"
        )
    if sim_sheet.get("visibility") != sim.layout.required_visibility:
        raise SimContractError(
            f"{sim.source_path}: sim_contract requires {sim.layout.sheet} visibility "
            f"{sim.layout.required_visibility!r}, but workbook.yaml declares "
            f"{sim_sheet.get('visibility')!r}. _SimData is machine data with no audit value in "
            "raw form; a visible or merely hidden sheet invites hand editing of retained samples."
        )

    results_sheet = _step(sheets, "Results")
    if results_sheet is _MISSING or not isinstance(results_sheet, dict):
        raise SimContractError(f"{sim.source_path}: workbook.yaml declares no Results sheet")
    titles = {
        block.get("title")
        for block in (results_sheet.get("blocks") or [])
        if isinstance(block, dict)
    }
    required_sections = tuple(sim.raw["results_minimum"]["sections"])
    missing_sections = [name for name in required_sections if name not in titles]
    if missing_sections:
        raise SimContractError(
            f"{sim.source_path}: the Results placeholder no longer declares section(s) "
            f"{missing_sections}, which results_minimum relies on. sim_contract must not depend "
            "on a placeholder that has moved."
        )

    model = workbook_doc.get("model") if isinstance(workbook_doc, dict) else None
    if not isinstance(model, dict) or not str(model.get("model_version") or "").strip():
        raise SimContractError(
            f"{sim.source_path}: workbook.yaml declares no model.model_version, but the run "
            "identity persists one"
        )

    # Distribution families: compare MEMBERSHIP, not order. The owner's list is
    # user-facing presentation order and has no reason to match dispatch order.
    input_doc = owners["input_contract.yaml"]
    master = _step(input_doc.get("config_tables"), "distributions")
    if master is _MISSING or not isinstance(master, dict):
        raise SimContractError(
            f"{sim.source_path}: input_contract.yaml declares no distributions master list"
        )
    owner_values = list(master.get("values") or [])
    if len(owner_values) != len(set(owner_values)):
        raise SimContractError(
            f"{sim.source_path}: the distributions master list contains a duplicate: "
            f"{owner_values}"
        )
    declared_families = list(sim.raw["distributions"]["families"])
    if set(declared_families) != set(owner_values):
        raise SimContractError(
            f"{sim.source_path}: the simulation families {sorted(declared_families)} and the "
            f"master list {sorted(owner_values)} disagree. Every accepted family must be "
            "simulable and no family may be simulable that the model does not offer."
        )

    # The selectable ladder is retained BY REFERENCE, so what is checked is that
    # the rule "store every selectable value" resolves against a real ladder.
    ladder = _step(input_doc.get("config_tables"), "confidence_levels")
    if ladder is _MISSING or not isinstance(ladder, dict):
        raise SimContractError(
            f"{sim.source_path}: input_contract.yaml declares no confidence-levels ladder"
        )
    ladder_values = list(ladder.get("values") or [])
    if not ladder_values:
        raise SimContractError(f"{sim.source_path}: the confidence-levels ladder is empty")
    if sim.raw["statistics"]["include_all_selectable_ladder_values"] is not True:
        raise SimContractError(
            f"{sim.source_path}: statistics must retain every selectable ladder value, so "
            "Selected Px is a deterministic lookup rather than a computation over samples"
        )
    fixed = list(sim.raw["statistics"]["fixed_nonselectable_percentiles"])
    overlap = set(fixed) & set(ladder_values)
    if overlap:
        raise SimContractError(
            f"{sim.source_path}: {sorted(overlap)} is both a fixed non-selectable percentile "
            "and a selectable ladder value"
        )

    # The one reference whose CONTENT this contract depends on: the seed domain
    # must actually be a whole-number BETWEEN rule that permits blank, or the
    # reference points at an authority that no longer says what is assumed here.
    seed = contract.inputs.get("random_seed")
    if seed is None:
        raise SimContractError(
            f"{sim.source_path}: input_contract.yaml declares no 'random_seed' input, so the "
            "seed-domain reference has no owner"
        )
    validation = seed.validation
    if not isinstance(validation, dict):
        raise SimContractError(
            f"{sim.source_path}: input_contract.yaml leaves random_seed.validation unset. "
            "D6-19 closed on input-contract ownership, and the deferred note is discharged by "
            "declaring the domain THERE, not by copying it here."
        )
    if validation.get("kind") != "whole" or validation.get("operator") != "between":
        raise SimContractError(
            f"{sim.source_path}: random_seed.validation must be a whole-number 'between' rule, "
            f"got kind={validation.get('kind')!r} operator={validation.get('operator')!r}"
        )
    if str(validation.get("formula1")) != "1" or str(validation.get("formula2")) != "2147483646":
        raise SimContractError(
            f"{sim.source_path}: random_seed.validation bounds are "
            f"{validation.get('formula1')!r}..{validation.get('formula2')!r}; the accepted "
            "D6-20 domain is 1..2147483646"
        )
    if validation.get("allow_blank") is False:
        raise SimContractError(
            f"{sim.source_path}: random_seed.validation forbids blank. Blank is not an "
            "omission here - it is the AUTO request."
        )
    if seed.required:
        raise SimContractError(
            f"{sim.source_path}: random_seed must remain optional; blank means AUTO"
        )


def retained_percentiles(sim: SimContract, contract_document: dict[str, Any]) -> tuple[str, ...]:
    """The percentiles a run must store, RESOLVED from the owning authority.

    The ladder's values live in `input_contract.yaml` and are not copied into the
    simulation contract, so this is how a caller - or a test - learns what the
    "store every selectable value" rule actually amounts to today.
    """
    ladder = _step(contract_document.get("config_tables"), "confidence_levels")
    if ladder is _MISSING or not isinstance(ladder, dict):
        raise SimContractError("input_contract.yaml declares no confidence-levels ladder")
    fixed = tuple(sim.raw["statistics"]["fixed_nonselectable_percentiles"])
    return fixed + tuple(ladder.get("values") or ())


_MISSING = object()


def _step(node: Any, part: str) -> Any:
    """One step of a dotted authority locator.

    A locator names a concept, not a YAML container type, so it resolves against a
    mapping key or against a list entry carrying its own identifier.
    """
    if isinstance(node, dict):
        return node.get(part, _MISSING)
    if isinstance(node, list):
        for item in node:
            if isinstance(item, dict):
                for id_key in ("key", "name", "concept"):
                    if item.get(id_key) == part:
                        return item
        return _MISSING
    return _MISSING


def _document_of(loaded: Any) -> Any:
    """The raw YAML document behind a loaded contract, however it retains it."""
    for attribute in ("raw", "document", "_raw"):
        value = getattr(loaded, attribute, None)
        if isinstance(value, dict):
            return value
    path = getattr(loaded, "source_path", None)
    if path is None:
        raise SimContractError(
            f"cannot resolve authority references against {type(loaded).__name__}: it exposes "
            "neither a raw document nor a source path"
        )
    return yaml.safe_load(Path(path).read_text(encoding="utf-8"))


# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------
def _walk(node: Any, prefix: tuple = ()):
    if isinstance(node, dict):
        for key, value in node.items():
            yield from _walk(value, prefix + (key,))
    elif isinstance(node, list):
        for i, value in enumerate(node):
            yield from _walk(value, prefix + (i,))
    else:
        yield prefix, node


def _flatten_text(node: Any) -> str:
    return " ".join(str(value) for _, value in _walk(node))


def _req(mapping: Any, key: str, where: str) -> Any:
    if not isinstance(mapping, dict) or key not in mapping:
        raise SimContractError(f"{where}: missing required key {key!r}")
    return mapping[key]


def _req_str(mapping: Any, key: str, where: str) -> str:
    value = _req(mapping, key, where)
    if not isinstance(value, str) or not value.strip():
        raise SimContractError(f"{where}: {key!r} must be a non-empty string, got {value!r}")
    return value


def _req_int(mapping: Any, key: str, where: str) -> int:
    value = _req(mapping, key, where)
    if not isinstance(value, int) or isinstance(value, bool):
        raise SimContractError(f"{where}: {key!r} must be an integer, got {value!r}")
    return value


def _map(mapping: Any, key: str, where: str) -> dict:
    value = _req(mapping, key, where)
    if not isinstance(value, dict):
        raise SimContractError(f"{where}: {key!r} must be a mapping, got {type(value).__name__}")
    return value


def _seq(mapping: Any, key: str, where: str) -> list:
    value = _req(mapping, key, where)
    if not isinstance(value, list) or not value:
        raise SimContractError(f"{where}: {key!r} must be a non-empty list")
    return value


def _require_value(mapping: Any, key: str, expected: Any, where: str) -> None:
    actual = _req(mapping, key, where)
    if actual != expected or isinstance(actual, bool) != isinstance(expected, bool):
        raise SimContractError(
            f"{where}: {key!r} must be {expected!r}, got {actual!r}. This value is settled "
            "authority; the contract encodes it and does not choose it."
        )


def _require_true(mapping: Any, key: str, where: str) -> None:
    if _req(mapping, key, where) is not True:
        raise SimContractError(f"{where}: {key!r} must be true")


def _require_false(mapping: Any, key: str, where: str) -> None:
    if _req(mapping, key, where) is not False:
        raise SimContractError(f"{where}: {key!r} must be false")


def _exact_sequence(actual: Any, expected: tuple, where: str) -> None:
    if not isinstance(actual, (list, tuple)):
        raise SimContractError(f"{where}: must be a list, got {actual!r}")
    if tuple(actual) != tuple(expected):
        raise SimContractError(
            f"{where}: must be exactly {list(expected)}, got {list(actual)}. Order and "
            "membership are both load-bearing."
        )
