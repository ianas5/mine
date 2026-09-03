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

LOCKED_REQUEST_RECORD_COUNT = 1
"""ONE SIM record. Five one-field records would encode the same semantics into
different bytes, and both would have satisfied the pre-closure contract."""

LOCKED_REQUEST_FIELD_TYPES = {
    "iterations": "F_I",
    "seed_mode": "F_S",
    "supplied_seed": "F_I",
    "rng_version": "F_I",
    "sim_method_version": "F_I",
}
"""One canonical encoder per field. F_I for every integer identity: a count, a
seed and a version are STRUCTURAL facts, and F_N would let a version of 1 encode
identically to a Double of 1."""

LOCKED_REQUEST_AUTO_FIELDS = ("iterations", "seed_mode", "rng_version", "sim_method_version")
LOCKED_REQUEST_FIXED_FIELDS = (
    "iterations", "seed_mode", "supplied_seed", "rng_version", "sim_method_version"
)
LOCKED_REQUEST_EFFECTIVE = {
    "AUTO": LOCKED_REQUEST_AUTO_FIELDS,
    "FIXED": LOCKED_REQUEST_FIXED_FIELDS,
}
"""AUTO is FOUR fields because the supplied seed does not exist there. Zero,
blank, null and the previous effective seed are all different streams, and each
would break the recomputability an AUTO request depends on."""

LOCKED_REQUEST_GRAMMAR = {
    "section": 'F_S("SIM") F_I(1) sim_record',
    "auto_record": (
        'F_I(4) F_I(iterations) F_S("AUTO") F_I(rng_version) F_I(sim_method_version)'
    ),
    "fixed_record": (
        'F_I(5) F_I(iterations) F_S("FIXED") F_I(supplied_seed) F_I(rng_version) '
        'F_I(sim_method_version)'
    ),
}
"""The stream, token by token - the same standard `result_digest` has carried
since Step 0. A grammar the validator does not check is a grammar the
implementation gets to choose."""

LOCKED_REQUEST_SEED_ABSENT = "absent"
LOCKED_REQUEST_SEED_DOMAIN_OWNER = "input_contract.yaml"
LOCKED_REQUEST_STREAM_TAG_OWNER = "calc_contract.yaml"

LOCKED_SIM_STATES = ("CURRENT", "STALE", "INVALID")
# PHASE-6 ONLY, AND A STRICT SUPERSET OF THE PHASE-5 AXIS. Phase 5 keeps
# exactly ("NONE", "SUCCESS", "REFUSED", "FAILED") and is locked separately in
# calc_loader.py; the fifth token names a persistence/recovery condition that
# has no Phase-5 analogue. Still an EXACT sequence: order and membership are
# both load-bearing, and this is an authority correction, not a relaxation.
LOCKED_ATTEMPT_RESULTS = (
    "NONE", "SUCCESS", "REFUSED", "FAILED", "AUTO_NONCE_INDETERMINATE",
)
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

LOCKED_RECURRENCE_ADVANCE = '[s10, s11, s12, s20, s21, s22] <- [s11, s12, p1, s21, s22, p2]'
"""The state shift. Part of the exact MRG32k3a contract, not documentation: a
recurrence that keeps the right p1/p2 but shifts the wrong words produces a
plausible stream that is not MRG32k3a."""

LOCKED_BC_PER_DRIVER = (
    'alpha = a + b',
    'beta = 1 / b',
    'delta = 1 + a - b',
    'k1 = delta * (0.0138889 + 0.0416667 * b) / (a * beta - 0.777778)',
    'k2 = 0.25 + (0.5 + 0.25 / delta) * b',
)
LOCKED_BC_PER_ATTEMPT = (
    'u1 = next_u(); u2 = next_u()',
    'if u1 < 0.5: y = u1 * u2; z = u1 * y; reject if 0.25 * u2 + z - y >= k1',
    'else: z = u1 * u1 * u2',
    '  if z <= 0.25: vlog = log(u1 / (1 - u1)); v = beta * vlog; w = a * exp(v); ACCEPT',
    '  if z >= k2: reject',
    'vlog = log(u1 / (1 - u1)); v = beta * vlog; w = a * exp(v)',
    'accept if alpha * (log(alpha / (b + w)) + v) - 1.3862944 >= log(z)',
    'else reject and retry',
)
LOCKED_BB_RETURN = "w / (b + w) when the caller's first parameter was the min, else b / (b + w)"
LOCKED_BC_RETURN = "w / (b + w) when the caller's first parameter was the max, else b / (b + w)"
"""BB and BC orient OPPOSITELY, so their return rules are opposite too. Locking
one and leaving the other free is how a mirrored distribution gets shipped."""

LOCKED_CHENG_FORMULATION_EVIDENCE = 'evidence/phase6_step0/raw/cheng_formulation.json'
LOCKED_CHENG_VECTORS_EVIDENCE = 'evidence/phase6_step0/vectors/cheng_vectors.json'
"""A contract that claims to bind evidence but may point anywhere is not bound."""

LOCKED_CHENG_SOURCE_SHA256 = "d5ca71b806015ed9039f713295dca3b45d2f66dbcabe60b6709896e3a89eed90"
LOCKED_A1_P127_SHA256 = "e31a727398a2d4461cf708f77034b9bc5e60f88c54556c56c3f4b015a813b66a"
LOCKED_A2_P127_SHA256 = "0d20b47aa206b1231c22e20afaa84e71b81cff52a2395be9aeb0bbb97b1e8208"
"""Authoritative-looking metadata the validator ignores is worse than none, so
these are checked against the accepted Step-0 values rather than merely
shape-checked. They are LITERALS here: production never reads `evidence/`."""

LOCKED_DIGEST_GRAMMAR = {
    "stream": 'F_S("PCCM-RD") F_I(sim_method_version) section',
    "section": 'F_S("RESULT") F_I(record_count) record*',
    "record": 'F_I(field_count) F_I(iteration_index) F_N(total_nominal) F_N(total_pv)',
}

LOCKED_CONDITIONING_SCALE = 's = max(abs(a), abs(m), abs(b))'
LOCKED_TRIANGULAR_BOUNDARY = {
    "m_equals_a": 'c = 0; the upper branch is always taken',
    "m_equals_b": 'c = 1; the lower branch is always taken',
}
LOCKED_DEGENERATE_CONDITIONS = {
    "uniform": "a == b",
    "triangular": "a == m == b",
    "beta_pert": "a == m == b",
}
"""FAMILY-SPECIFIC. Uniform's Most Likely is ignored by accepted Phase-5 D1, so a
common `a == m == b` predicate let an ignored input decide degeneracy - and
therefore RNG consumption and every later draw on that stream."""

LOCKED_RUN_IDENTITY = (
    ('last_successful_stamp', 8, 'snapshot', 'Last Successful Simulation', 'timestamp', None, None),
    ('run_id', 9, 'snapshot', 'Run ID', 'integer', None, None),
    ('request_fingerprint', 10, 'snapshot', 'Request Fingerprint', 'text', None, None),
    ('result_digest', 11, 'snapshot', 'Result Digest', 'text', None, None),
    ('seed_mode', 12, 'snapshot', 'Seed Mode', 'enum', 'seed_mode', None),
    ('supplied_seed', 13, 'snapshot', 'Supplied Seed', 'integer', None, None),
    ('effective_seed', 14, 'snapshot', 'Effective Seed', 'integer', None, None),
    ('consumed_auto_nonce', 15, 'snapshot', 'Consumed AUTO Nonce', 'integer', None, None),
    ('iterations_run', 16, 'snapshot', 'Iterations Run', 'integer', None, None),
    ('rng_version', 17, 'snapshot', 'RNG Version', 'integer', None, None),
    ('sim_method_version', 18, 'snapshot', 'Simulation Method Version', 'integer', None, None),
    ('model_version', 19, 'snapshot', 'Model Version', 'text', None, None),
    ('applied_timeline', 20, 'snapshot', 'Applied Timeline', 'text', None, None),
    ('next_auto_nonce', 21, 'counter', 'Next AUTO Nonce', 'integer', None, 0),
    ('last_run_id', 22, 'counter', 'Last Run ID', 'integer', None, 0),
    ('last_attempt_result', 23, 'attempt', 'Last Attempt Result', 'enum', 'attempt_result', 'NONE'),
    ('last_attempt_detail', 24, 'attempt', 'Last Attempt Detail', 'text', None, None),
    ('last_attempt_seed_mode', 25, 'attempt', 'Last Attempt Seed Mode', 'enum', 'seed_mode', None),
    ('last_attempt_effective_seed', 26, 'attempt', 'Last Attempt Effective Seed', 'integer', None, None),
    ('last_attempt_auto_nonce', 27, 'attempt', 'Last Attempt AUTO Nonce', 'integer', None, None),
    ('simulation_status', 28, 'derived', 'Simulation Status (last evaluated)', 'enum', 'sim_state', None),
    ('status_evaluated_at', 29, 'derived', 'Status Evaluated At', 'timestamp', None, None),
    ('active_bank', 30, 'control', 'Active Bank', 'enum', 'bank', None),
)
"""(key, row, group, label, value_type, enum, initial) - the COMPLETE record.

Key/row/group/type alone was not exact authority: initials could be seeded and
enum owners swapped, so a materialiser could have written a partial successful
snapshot into a workbook that had never run."""

LOCKED_RUN_IDENTITY_COLUMNS = {
    "label_column": 'B',
    "value_column": 'D',
    "note_column": 'H',
}
LOCKED_BANK_LABELS = ('A', 'B')
LOCKED_RUN_IDENTITY_BANK_COLUMNS = {'A': 'D', 'B': 'F'}
'''The BANKED snapshot columns. Bank A reuses the shared value column, which is
what keeps a first-ever successful run in exactly the place the accepted
single-bank layout put it.'''

LOCKED_ITERATION_BANKS = {
    'A': {'iteration_index': 'B', 'total_nominal': 'C', 'total_pv': 'D'},
    'B': {'iteration_index': 'F', 'total_nominal': 'G', 'total_pv': 'H'},
}
LOCKED_CANDIDATE_TARGET = {'': 'A', 'A': 'B', 'B': 'A'}

LOCKED_TRANSACTION_ORDER = (
    'prepare_phase5_inputs_and_require_current',
    'validate_pre_allocation_prerequisites',
    'allocate_auto_nonce_when_auto',
    'run_simulation_and_statistics_in_memory',
    'build_request_fingerprint_and_result_digest_in_memory',
    'choose_inactive_bank',
    'write_candidate_snapshot_to_inactive_bank',
    'write_candidate_summary_to_inactive_bank',
    'write_candidate_contingency_to_inactive_bank',
    'write_candidate_iterations_to_inactive_bank',
    'verify_inactive_bank_against_staged_package',
    'final_commit_shared_block_including_active_bank',
)
LOCKED_FINAL_COMMIT_RANGE = 'D22:D30'
LOCKED_FINAL_COMMIT_FIELDS = (
    'last_run_id', 'last_attempt_result', 'last_attempt_detail',
    'last_attempt_seed_mode', 'last_attempt_effective_seed', 'last_attempt_auto_nonce',
    'simulation_status', 'status_evaluated_at', 'active_bank',
)
'''The shared block, CONTIGUOUS and in row order, ending with the publication
selector. One Range write, and the active bank moves last inside it.'''

LOCKED_READ_ACCESSORS = (
    'PCCM_SimulationStatus',
    'PCCM_SimulationRequestFingerprint',
    'PCCM_CurrentSimulationRequestFingerprint',
    'PCCM_SimulationResultDigest',
    'PCCM_SimulationAttemptResult',
    'PCCM_SimulationAttemptDetail',
)

LOCKED_READ_ACCESSOR_SEMANTICS = {
    'PCCM_SimulationStatus':
        'the derived status; writes only simulation_status and status_evaluated_at',
    'PCCM_SimulationRequestFingerprint':
        'the stored request fingerprint of the ACTIVE bank, blank when no bank is active',
    'PCCM_CurrentSimulationRequestFingerprint':
        'the request fingerprint recomputed through the SAME preparation path the run '
        'uses, blank when the current prerequisites refuse',
    'PCCM_SimulationResultDigest':
        'the stored result digest of the ACTIVE bank, blank when no bank is active',
    'PCCM_SimulationAttemptResult': 'the shared last attempt result',
    'PCCM_SimulationAttemptDetail': 'the shared last attempt detail',
}
'''Exact wording, for the same reason the three simulation-state definitions are
exact: `SimulationRequestFingerprint` reads a stored value and
`CurrentSimulationRequestFingerprint` recomputes one, and a prose drift between
them is the difference between a correct staleness answer and a wrong one.'''

LOCKED_SUMMARY_METRICS = (
    # (key, row, label, source). The QUANTILE labels are deliberately None: the
    # selectable ladder belongs to input_contract.yaml, and spelling "P55" here
    # would be a second ladder that could drift from the owning authority.
    ('mean', 8, 'Mean', 'SimStatsMean'),
    ('sample_standard_deviation', 9, 'Sample Standard Deviation',
     'SimStatsSampleStandardDeviation'),
    ('minimum', 10, 'Minimum', 'SimStatsDescribe'),
    ('quantile_1', 11, None, 'SimStatsDescribe'),
    ('quantile_2', 12, None, 'SimStatsDescribe'),
    ('quantile_3', 13, None, 'SimStatsDescribe'),
    ('quantile_4', 14, None, 'SimStatsDescribe'),
    ('quantile_5', 15, None, 'SimStatsDescribe'),
    ('quantile_6', 16, None, 'SimStatsDescribe'),
    ('quantile_7', 17, None, 'SimStatsDescribe'),
    ('quantile_8', 18, None, 'SimStatsDescribe'),
    ('quantile_9', 19, None, 'SimStatsDescribe'),
    ('quantile_10', 20, None, 'SimStatsDescribe'),
    ('quantile_11', 21, None, 'SimStatsDescribe'),
    ('maximum', 22, 'Maximum', 'SimStatsDescribe'),
    ('deterministic_base_a', 23, 'Deterministic Base A', 'phase5_preparation'),
)
LOCKED_SUMMARY_COLUMNS = {
    'label_column': 'J',
    'A': {'nominal': 'K', 'pv': 'L'},
    'B': {'nominal': 'M', 'pv': 'N'},
}
LOCKED_CONTINGENCY_COLUMNS = {
    'label_column': 'P',
    'A': {'nominal': 'Q', 'pv': 'R'},
    'B': {'nominal': 'S', 'pv': 'T'},
}
LOCKED_RESULTS_RUN_STAMP_FIELDS = (
    'last_successful_stamp', 'run_id', 'model_version', 'iterations_run', 'seed_mode',
    'supplied_seed', 'effective_seed', 'consumed_auto_nonce', 'applied_timeline',
    'rng_version', 'sim_method_version', 'request_fingerprint', 'result_digest',
    'simulation_status', 'status_evaluated_at',
)
LOCKED_RESULTS_FORBIDDEN_FUNCTIONS = (
    'AVERAGE', 'STDEV', 'STDEV.S', 'STDEVP', 'PERCENTILE', 'PERCENTILE.INC',
    'PERCENTILE.EXC', 'QUARTILE', 'MEDIAN', 'OFFSET', 'INDIRECT', 'RAND',
    'RANDBETWEEN', 'NOW', 'TODAY',
)
LOCKED_PHASE5_BRIDGE_RETURNS = (
    'drivers', 'driver_count', 'analytical_fingerprint', 'deterministic_base_nominal',
    'deterministic_base_pv', 'applied_timeline', 'decimal_separator',
)

LOCKED_RESERVED_ROWS = (
    (1, 1, 'shell top margin'),
    (2, 2, 'shell title'),
    (3, 3, 'shell subtitle'),
    (4, 4, 'shell rule'),
    (5, 5, 'spacer'),
    (6, 6, 'run identity section heading'),
    (7, 7, 'run identity section note'),
    (8, 30, 'run identity fields'),
    (31, 31, 'iteration records section heading'),
    (32, 32, 'iteration records section note'),
    (33, 33, 'iteration table header'),
)

LOCKED_ITERATION_RECORD_COLUMNS = (
    ('iteration_index', 'B', 'Iteration', 'integer'),
    ('total_nominal', 'C', 'Total Nominal', 'double'),
    ('total_pv', 'D', 'Total PV', 'double'),
)

LOCKED_SIM_STATE_DEFINITIONS = {
    'CURRENT': 'prerequisites resolve, a successful snapshot exists, and the recomputed request fingerprint equals the stored successful one',
    'STALE': 'prerequisites resolve, a successful snapshot exists, and the recomputed request fingerprint differs from the stored successful one',
    'INVALID': 'current simulation prerequisites do not resolve',
}

LOCKED_INDEX_RULE = 'stream k is the base state advanced by k canonical stream jumps'
LOCKED_RNG_VERSION = 1
LOCKED_SIM_METHOD_VERSION = 1
"""Settled by Step 0. A future bump is an explicit authority change, not
something a validator should wave through because 2 is also a positive integer."""

LOCKED_SIM_DATA_SHEET = "_SimData"
LOCKED_ANNUAL_QUANTILE_COUNT = 11

# The Phase-7 column allocation, locked. Letters are the persistence layout, so
# a silent change to one of them moves published data.
LOCKED_SENSITIVITY_COLUMN_LAYOUT = (
    ("driver_id", "J", "Driver ID", "text"),
    ("driver_type", "K", "Type", "text"),
    ("driver_name", "L", "Name", "text"),
    ("rho", "M", "Rho", "double"),
    ("abs_rho", "N", "|Rho|", "double"),
    ("rank", "O", "Rank", "integer"),
    ("direction", "P", "Direction", "text"),
    ("status", "Q", "Status", "text"),
)
LOCKED_SENSITIVITY_ROW_RULE = "one row per eligible driver, in ranked order"

# THE STAMP, LOCKED. `published` is last on purpose: it is written after the
# rows and the identity, so a write that fails part way leaves the block
# unpublished rather than current.
LOCKED_SENSITIVITY_STAMP = (
    ("run_id", 8, "integer"),
    ("effective_seed", 9, "integer"),
    ("request_fingerprint", 10, "text"),
    ("result_digest", 11, "text"),
    ("iterations", 12, "integer"),
    ("record_count", 13, "integer"),
    ("published", 14, "text"),
)
LOCKED_SENSITIVITY_STAMP_COLUMNS = {"A": "J", "B": "S"}
LOCKED_ANNUAL_ROW_RULE = "one row per applied project year"
LOCKED_ANNUAL_INDEX_COLUMNS = {
    "A": {"project_index": "AB", "calendar_year": "AC"},
    "B": {"project_index": "BC", "calendar_year": "BD"},
}
LOCKED_ANNUAL_QUANTILE_FIRST_COLUMN = {
    "A": {"nominal": "AD", "pv": "AO"},
    "B": {"nominal": "BE", "pv": "BP"},
}
LOCKED_ANNUAL_PROFILE_COLUMNS = {
    "A": {"nominal": "AZ", "pv": "BA"},
    "B": {"nominal": "CA", "pv": "CB"},
}
LOCKED_ITERATION_COLUMNS = ("iteration_index", "total_nominal", "total_pv")
LOCKED_SIM_STATE_RULES = (
    (1, "current_prerequisites_do_not_resolve", "INVALID"),
    (2, "no_successful_snapshot_exists", None),
    (3, "request_fingerprint_equals_stored_successful", "CURRENT"),
    (4, "request_fingerprint_differs_from_stored_successful", "STALE"),
)
"""The corrected derivation - ordered, total, and blind to the attempt history."""

LOCKED_SIM_DATA_EXCLUDED = (
    "per_driver_samples",
    "annual_iteration_matrix",
)
"""The retentions that stay refused.

PHASE 7 CHANGED THIS LIST, AND THE CHANGE IS NARROWER THAN IT LOOKS.
`sensitivity_data` and `annual_stochastic_samples` left the list because Phase 7
CONTRACTS those outputs. What did not change is the memory architecture: the
driver x iteration matrix is still refused, and the iteration x year matrix is
refused under the name `annual_iteration_matrix`. Sensitivity reaches per-driver
values by reset-and-replay and annual output by block replay, precisely so that
neither matrix is ever created. Contracting an OUTPUT is not retaining the
MATRIX behind it.
"""

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
    "sensitivity",
    "annual_stochastic",
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
        'accumulation', 'annual_stochastic', 'authority_references', 'cheng',
        'command_surface', 'components',
        'contingency', 'contribution', 'dependence', 'distributions', 'interruption',
        'iterations', 'jump', 'kernel', 'label_sets', 'numerical_domain', 'prerequisite',
        'phase5_bridge', 'publication', 'request_fingerprint', 'result_digest',
        'results_minimum', 'risk', 'rng', 'run_id', 'seeding', 'selected_confidence_level',
        'sensitivity', 'sim_contract_version', 'sim_data', 'sim_state',
        'statistics', 'stream_assignment', 'versions'
    }),
    # ---------------------------------------------------------------------
    # PHASE 7. Sensitivity and annual stochastic output.
    # ---------------------------------------------------------------------
    'sensitivity': frozenset({
        'basis_run', 'contribution', 'display', 'drivers', 'identity_binding',
        'independent_monte_carlo_permitted', 'interpretation', 'kind', 'phase',
        'ranking', 'replay', 'sampling', 'state_safety', 'statistic', 'zero_variance'
    }),
    'sensitivity.drivers': frozenset({
        'category_aggregation', 'identity', 'identity_is_worksheet_row',
        'one_per_cost_line', 'one_per_risk'
    }),
    'sensitivity.contribution': frozenset({
        'correlation_against_raw_severity_permitted', 'cost_line', 'expression_owner',
        'measure', 'reimplementation_permitted', 'risk',
        'risk_occurrence_and_severity_are_one_driver'
    }),
    'sensitivity.replay': frozenset({
        'concurrent_driver_columns_retained', 'cost_line_streams', 'granularity',
        'random_access_seek', 'resets_from', 'retains_driver_matrix', 'risk_streams',
        'risk_streams_paired_by_iteration', 'sequential_advance_required',
        'unrelated_drivers_advanced'
    }),
    'sensitivity.statistic': frozenset({
        'definition', 'iteration_correspondence_preserved', 'name',
        'no_ties_shortcut_permitted', 'sorting', 'source_arrays_mutated', 'tie_rule',
        'total_ranks_computed_once', 'total_ranks_reused_across_drivers'
    }),
    'sensitivity.zero_variance': frozenset({
        'excluded_from_ranking', 'excluded_from_tornado_input', 'reported_as_zero_rho',
        'retained_diagnostically', 'rho_reported', 'status_label'
    }),
    'sensitivity.ranking': frozenset({
        'direction', 'direction_labels', 'order_by', 'population', 'signed_rho_retained',
        'tie_break', 'tie_break_comparison', 'tie_break_direction',
        'tie_break_uses_supply_order', 'tie_break_uses_worksheet_row',
        'top_n_truncation'
    }),
    'sensitivity.sampling': frozenset({
        'index_set_rule_if_ever_adopted', 'iteration_cap', 'sensitivity_sample_size',
        'subsampling_contracted', 'unmeasured_performance_safeguard_permitted',
        'uses_all_iterations'
    }),
    'sensitivity.sampling.index_set_rule_if_ever_adopted': frozenset({
        'full_run_statistics_recomputed_from_subset',
        'independently_selected_samples_permitted', 'same_indices', 'same_order',
        'shared_index_set'
    }),
    'sensitivity.interpretation': frozenset({
        'inter_driver_correlation_owner', 'measures', 'measures_variance_contribution',
        'percentage_contribution_permitted', 'rho_squared_as_variance_share_permitted'
    }),
    'sensitivity.identity_binding': frozenset({
        'stamped_with', 'stored_in', 'valid_only_for_the_stamped_run'
    }),
    'sensitivity.state_safety': frozenset({
        'advances_auto_nonce', 'changes_result_digest', 'consumes_run_id',
        'mutates_successful_snapshot', 'rewrites_iteration_records',
        'touches_pending_auto_nonce_marker', 'writes_attempt_row'
    }),
    'sensitivity.display': frozenset({
        'presented_as_current_when', 'presented_as_current_when_invalid',
        'presented_as_current_when_stale', 'prior_sensitivity_preserved_on_failure',
        'refused_attempt_destroys_prior_sensitivity', 'stale_must_be_labelled',
        'state_owner'
    }),
    'annual_stochastic': frozenset({
        'annual_distributions', 'contracted', 'iteration_annual_vector',
        'per_year_factor', 'phase', 'phase_8_handoff', 'retention',
        'selected_px_profile'
    }),
    'annual_stochastic.per_year_factor': frozenset({
        'decomposition_of', 'factor_owner', 'is_a_new_input', 'nominal', 'pv',
        'reconciles'
    }),
    'annual_stochastic.iteration_annual_vector': frozenset({
        'identity', 'identity_owner', 'nominal', 'pv'
    }),
    'annual_stochastic.retention': frozenset({
        'block_axis', 'block_width_configurable', 'passes',
        'persisted_iteration_by_year_matrix', 'retained', 'retained_in_memory_matrix',
        'strategy'
    }),
    'annual_stochastic.annual_distributions': frozenset({
        'is_a_selected_px_profile', 'ladder_owner', 'measures', 'method', 'per_year',
        'quantile_method_owner', 'sorting', 'sums_to_total_percentile'
    }),
    'annual_stochastic.selected_px_profile': frozenset({
        'definition', 'degenerates_to_single_iteration_when_f_is_zero', 'formula',
        'lo_hi_f_source', 'measures', 'nearest_rank_permitted',
        'other_definitions_permitted', 'per_year_percentile_as_profile_permitted',
        'position_owner', 'reconciliation_identity', 'reconciliation_rule_owner'
    }),
    'annual_stochastic.phase_8_handoff': frozenset({
        'annual_cash_flow_presentation_owner', 'chart_owner', 'dashboard_owner',
        'presentation_in_phase_7', 'provides'
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
        'automation_endpoint', 'effective_seed_public_accessor_required_in_phase_6',
        'msgbox_introduced_by_phase_6', 'read_accessor_names_settled', 'read_accessors',
        'read_accessor_semantics', 'ribbon_introduced_by_phase_6',
        'run_id_public_accessor_required_in_phase_6', 'user_facing_run_button_in_phase_6',
        'userform_introduced_by_phase_6'
    }),
    'command_surface.read_accessor_semantics': frozenset(LOCKED_READ_ACCESSORS),
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
        'applies_to_all_families', 'conditions', 'detected_before_dispatch',
        'detected_before_parameterisation', 'most_likely_read_by_uniform_degeneracy',
        'returns', 'sampler_entered', 'stream_state_changed', 'uniforms_consumed'
    }),
    'distributions.degenerate.conditions': frozenset({'beta_pert', 'triangular', 'uniform'}),
    'distributions.triangular': frozenset({
        'boundary_cases', 'branch_point', 'conditioning_scale', 'lower_branch', 'method',
        'normalised_formulation_required', 'rng_endpoints_open',
        'uniforms_per_non_degenerate_sample', 'upper_branch'
    }),
    'distributions.triangular.boundary_cases': frozenset({'m_equals_a', 'm_equals_b'}),
    'distributions.uniform': frozenset({
        'formulation', 'most_likely_affects_degeneracy',
        'most_likely_affects_uniform_consumption', 'most_likely_used', 'transform',
        'uniforms_per_non_degenerate_sample'
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
    'label_sets': frozenset({'attempt_result', 'bank', 'seed_mode', 'sim_state'}),
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
        'banks', 'commit_last', 'failure_semantics',
        'partial_new_distribution_published_on_refusal_or_failure',
        'persisted_source_of_truth', 'prior_successful_publication_survives',
        'publish_only_after_simulation_and_statistics_complete', 'results_derives_from',
        'results_recomputes_monte_carlo', 'run_id_allocation', 'transaction'
    }),
    'publication.banks': frozenset({
        'candidate_target', 'candidate_writes_to_active_bank', 'count',
        'duplicate_workbook_required', 'inactive_bank_is_published',
        'inactive_bank_is_staging_storage', 'initial_active_bank', 'labels',
        'row_axis_shared_by_both_banks', 'temporary_worksheet_required',
        'third_bank_permitted'
    }),
    'publication.banks.candidate_target': frozenset({'', 'A', 'B'}),
    'publication.transaction': frozenset({
        'final_commit_failure_restores_prior_block', 'final_commit_fields',
        'final_commit_is_one_write', 'final_commit_range',
        'million_row_rollback_required', 'order',
        'prior_final_commit_block_captured_before_write',
        'results_is_a_written_transaction'
    }),
    'publication.run_id_allocation': frozenset({
        'allocated_by', 'candidate_value', 'headroom_checked_before_auto_allocation',
        'held_locally_until_commit'
    }),
    'publication.failure_semantics': frozenset({
        'final_commit_failure', 'inactive_bank_write_failure',
        'refusal_before_auto_allocation', 'refusal_or_failure_after_auto_allocation'
    }),
    'publication.failure_semantics.refusal_before_auto_allocation': frozenset({
        'active_bank_changed', 'attempt_metadata_updated', 'next_auto_nonce_advanced',
        'successful_banks_changed'
    }),
    'publication.failure_semantics.refusal_or_failure_after_auto_allocation': frozenset({
        'active_bank_changed', 'attempt_metadata_updated', 'next_auto_nonce_advanced',
        'successful_banks_changed'
    }),
    'publication.failure_semantics.inactive_bank_write_failure': frozenset({
        'active_bank_changed', 'corrupted_candidate_has_semantic_standing',
        'prior_publication_remains_authoritative'
    }),
    'publication.failure_semantics.final_commit_failure': frozenset({
        'active_bank_changed', 'prior_block_restored'
    }),
    'phase5_bridge': frozenset({
        'analytical_fingerprint_is_current_not_stored', 'duplicates_factor_mathematics',
        'is_automation_endpoint', 'name_prefix_pccm', 'owner_module', 'procedure',
        'requires_phase5_status', 'returns', 'reuses_private_preparation',
        'updates_phase5_status_or_attempt_metadata', 'writes_to_calc_sheet',
        'zero_driver_model_succeeds'
    }),
    'selected_confidence_level': frozenset({
        'change_requires_rerun', 'invalid_selector_blanks_selected_reporting_rows',
        'invalid_selector_invalidates_simulation', 'participates_in_auto_allocation',
        'participates_in_execution_validity', 'participates_in_request_fingerprint',
        'participates_in_state_derivation', 'source', 'unselected_state_introduced'
    }),
    'request_fingerprint': frozenset({
        'analytical_prefix', 'auto_blank_seed_remains_recomputable',
        'existing_sections_modified', 'extension_semantics', 'section_order', 'sim_section'
    }),
    'request_fingerprint.sim_section': frozenset({
        'analytical_fingerprint_hashed_as_a_field', 'auto_supplied_seed_representation',
        'effective_records', 'encoded_field_names', 'excluded_fields', 'field_types', 'fields',
        'grammar', 'name', 'record_count', 'stream_tag_owner',
        'stream_tag_repeated_in_extension', 'stream_version_repeated_in_extension',
        'supplied_seed_domain_owner', 'supplied_seed_present_only_when'
    }),
    'request_fingerprint.sim_section.effective_records': frozenset({'AUTO', 'FIXED'}),
    'request_fingerprint.sim_section.effective_records.AUTO': frozenset({
        'field_count', 'fields'
    }),
    'request_fingerprint.sim_section.effective_records.FIXED': frozenset({
        'field_count', 'fields'
    }),
    'request_fingerprint.sim_section.field_types': frozenset({
        'iterations', 'rng_version', 'seed_mode', 'sim_method_version', 'supplied_seed'
    }),
    'request_fingerprint.sim_section.grammar': frozenset({
        'auto_record', 'fixed_record', 'section'
    }),
    'result_digest': frozenset({
        'equality', 'field_types', 'grammar', 'iteration_index_origin', 'order_source',
        'record_field_count', 'record_fields', 'samples_sorted_for_digest', 'section_name',
        'stream_tag', 'version_field_source'
    }),
    'result_digest.grammar': frozenset({'record', 'section', 'stream'}),
    'results_minimum': frozenset({
        'annual_simulated_samples_contracted', 'deferred', 'presentation', 'sections'
    }),
    'results_minimum.presentation': frozenset({
        'blank_when_no_active_bank', 'blank_when_selector_not_selectable',
        'computes_contingency', 'computes_statistics',
        'contingency_by_subtraction_on_results', 'forbidden_functions',
        'materialised_by_stage_a', 'reads_a_fixed_bank', 'reads_only',
        'recomputes_quantiles', 'run_stamp_fields', 'selected_rows',
        'summary_metrics_source', 'written_by_the_run'
    }),
    'results_minimum.presentation.selected_rows[]': frozenset({
        'key', 'lookup_only', 'source'
    }),
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
        'allocation_state_semantics', 'allocation_states', 'attempt_result_token',
        'attempt_metadata_preserves', 'exhausted_value', 'fixed_mode',
        'failure_after_allocation_consumes_nonce', 'failure_before_allocation_consumes_nonce',
        'first_valid_allocation', 'immediate_reconciliation',
        'indeterminate_marker_storage_failure', 'initial', 'last_valid_allocation', 'meaning',
        'next_run_reconciliation', 'on_exhaustion', 'order', 'pending_clear',
        'prior_successful_publication_untouched', 'reuse_permitted', 'wrap_permitted',
        'recovery_action', 'write_ahead_order',
        'write_ahead_order_refinement_drops_prefix', 'write_ahead_order_refines'
    }),
    'seeding.nonce_lifecycle.fixed_mode': frozenset({
        'clears_pending_auto_nonce', 'executes_after_nonce_failpoint',
        'may_overwrite_prior_auto_attempt_metadata',
        'may_proceed_while_pending_marker_exists', 'reads_next_auto_nonce',
        'reads_pending_auto_nonce', 'writes_next_auto_nonce', 'writes_pending_auto_nonce'
    }),
    'seeding.nonce_lifecycle.attempt_result_token': frozenset({
        'audit_only', 'is_the_durable_recovery_authority',
        'may_be_overwritten_by_a_later_attempt', 'other_unsuccessful_outcomes_record',
        'phase5_axis_unchanged', 'recorded_for', 'recorded_for_recovery_action', 'token'
    }),
    'seeding.nonce_lifecycle.recovery_action': frozenset({
        'allocation_state_when_raised_before_identity_selection',
        'allocation_state_when_raised_by_cleanup',
        'carried_separately_from_allocation_state', 'derives_physical_consumption',
        'is_an_allocation_state', 'may_revise_allocation_state', 'permits_sampling',
        'raised_by', 'token'
    }),
    'seeding.nonce_lifecycle.pending_clear': frozenset({
        'counter_rollback_on_clear_failure', 'is_a_real_com_write',
        'raised_clear_proves_marker_remains', 'unresolved_cleanup_permits_sampling'
    }),
    'seeding.nonce_lifecycle.allocation_state_semantics': frozenset({
        'CONSUMED', 'PERSISTENCE_INDETERMINATE', 'PRE_ALLOCATION'
    }),
    'seeding.nonce_lifecycle.allocation_state_semantics.CONSUMED': frozenset({
        'advance_persisted', 'nonce_consumed', 'retry_may_reuse_the_same_nonce',
        'sampling_may_begin'
    }),
    'seeding.nonce_lifecycle.allocation_state_semantics.PRE_ALLOCATION': frozenset({
        'advance_persisted', 'nonce_consumed', 'retry_may_reuse_the_same_nonce',
        'sampling_began'
    }),
    'seeding.nonce_lifecycle.allocation_state_semantics.PERSISTENCE_INDETERMINATE': frozenset({
        'advance_persisted', 'audit_attempt_result',
        'audit_result_is_the_recovery_authority', 'durable_recovery_authority',
        'must_not_be_called_allocated', 'must_not_be_called_unconsumed',
        'next_run_must_reconcile_before_allocating', 'nonce_consumed', 'sampling_began'
    }),
    'seeding.nonce_lifecycle.immediate_reconciliation': frozenset({
        'after', 'attempts', 'observation_unavailable', 'observed_m',
        'observed_m_plus_1', 'observed_other'
    }),
    'sim_data.pending_auto_nonce': frozenset({
        'blank_means', 'cell', 'cleared_on', 'column',
        'counter_persist_forbidden_until_established',
        'integer_means', 'is_a_consumed_nonce_claim',
        'is_an_enum_sentinel', 'label', 'retained_on', 'row',
        'sampling_forbidden_while_unresolved', 'survives_unrelated_attempts',
        'value_type', 'written_before_counter_persist'
    }),
    'seeding.nonce_lifecycle.next_run_reconciliation': frozenset({
        'activated_by_attempt_result', 'activated_by_generic_unsuccessful_result',
        'activated_by_pending_auto_nonce_cell',
        'applies_to_fixed_mode', 'counter_equals_m', 'counter_equals_m_plus_1',
        'counter_is_neither', 'counter_unreadable'
    }),
    'seeding.nonce_lifecycle.attempt_metadata_preserves': frozenset({
        'known_consumed', 'persistence_indeterminate', 'pre_allocation'
    }),
    'seeding.nonce_lifecycle.indeterminate_marker_storage_failure': frozenset({
        'attempt_row_failure_loses_audit_line_only',
        'reuse_prevention_depends_on_attempt_row',
        'reuse_prevention_depends_on_pending_cell', 'second_write_ahead_log_required'
    }),
    'seeding.scalar_to_state': frozenset({
        'alternate_expansion_permitted', 'expansion', 'mixer', 'rule'
    }),
    'sim_data.sensitivity_records': frozenset({
        'banks', 'columns', 'consumes_reserved_rows', 'first_record_row', 'footer_rows',
        'header_row', 'row_rule', 'shares_row_axis_with_iteration_records', 'stamp'
    }),
    'sim_data.sensitivity_records.stamp': frozenset({
        'bank_value_columns', 'cleared_before_write', 'fields',
        'published_written_last', 'surplus_rows_cleared'
    }),
    'sim_data.sensitivity_records.stamp.bank_value_columns': frozenset({'A', 'B'}),
    'sim_data.sensitivity_records.stamp.fields[]': frozenset({'key', 'row', 'value_type'}),
    'sim_data.sensitivity_records.columns[]': frozenset({
        'column', 'header', 'key', 'value_type'
    }),
    'sim_data.sensitivity_records.banks': frozenset({'A', 'B'}),
    'sim_data.sensitivity_records.banks.A': frozenset({'first_column', 'last_column'}),
    'sim_data.sensitivity_records.banks.B': frozenset({'first_column', 'last_column'}),
    'sim_data.annual_records': frozenset({
        'consumes_reserved_rows', 'first_record_row', 'footer_rows', 'header_row',
        'index_columns', 'iteration_level_annual_values_persisted', 'max_rows_owner',
        'quantile_count', 'quantile_first_column', 'quantile_keys_owner', 'row_rule',
        'selected_px_profile_columns', 'shares_row_axis_with_iteration_records'
    }),
    'sim_data.annual_records.index_columns': frozenset({'A', 'B'}),
    'sim_data.annual_records.index_columns.A': frozenset({'calendar_year', 'project_index'}),
    'sim_data.annual_records.index_columns.B': frozenset({'calendar_year', 'project_index'}),
    'sim_data.annual_records.quantile_first_column': frozenset({'A', 'B'}),
    'sim_data.annual_records.quantile_first_column.A': frozenset({'nominal', 'pv'}),
    'sim_data.annual_records.quantile_first_column.B': frozenset({'nominal', 'pv'}),
    'sim_data.annual_records.selected_px_profile_columns': frozenset({'A', 'B'}),
    'sim_data.annual_records.selected_px_profile_columns.A': frozenset({'nominal', 'pv'}),
    'sim_data.annual_records.selected_px_profile_columns.B': frozenset({'nominal', 'pv'}),
    'sim_data': frozenset({
        'annual_records', 'contingency_ladder', 'excluded', 'iteration_records',
        'pending_auto_nonce', 'sensitivity_records',
        'required_visibility', 'reserved_rows', 'run_identity', 'sheet',
        'summary_statistics'
    }),
    'sim_data.contingency_ladder': frozenset({
        'all_values_representable_required_before_commit', 'bank_value_columns',
        'baseline', 'computed_for_whole_ladder_before_commit',
        'fixed_rung_persisted_though_not_selectable', 'first_row', 'label_column',
        'last_row', 'rungs', 'source', 'worksheet_subtraction_permitted'
    }),
    'sim_data.contingency_ladder.bank_value_columns': frozenset({'A', 'B'}),
    'sim_data.contingency_ladder.bank_value_columns.A': frozenset({'nominal', 'pv'}),
    'sim_data.contingency_ladder.bank_value_columns.B': frozenset({'nominal', 'pv'}),
    'sim_data.contingency_ladder.rungs[]': frozenset({'key', 'label', 'row'}),
    'sim_data.summary_statistics': frozenset({
        'bank_value_columns', 'first_row', 'label_column', 'last_row', 'metrics',
        'recomputed_from_worksheet_data', 'source'
    }),
    'sim_data.summary_statistics.bank_value_columns': frozenset({'A', 'B'}),
    'sim_data.summary_statistics.bank_value_columns.A': frozenset({'nominal', 'pv'}),
    'sim_data.summary_statistics.bank_value_columns.B': frozenset({'nominal', 'pv'}),
    'sim_data.summary_statistics.metrics[]': frozenset({'key', 'label', 'row', 'source'}),
    'sim_data.iteration_records': frozenset({
        'banks', 'columns', 'first_iteration_row', 'footer_rows', 'header_row', 'order',
        'sorted'
    }),
    'sim_data.iteration_records.banks': frozenset({'A', 'B'}),
    'sim_data.iteration_records.banks.A': frozenset({
        'iteration_index', 'total_nominal', 'total_pv'
    }),
    'sim_data.iteration_records.banks.B': frozenset({
        'iteration_index', 'total_nominal', 'total_pv'
    }),
    'sim_data.iteration_records.columns[]': frozenset({'column', 'header', 'key', 'value_type'}),
    'sim_data.reserved_rows[]': frozenset({'purpose', 'rows'}),
    'sim_data.run_identity': frozenset({
        'bank_value_columns', 'fields', 'first_row', 'label_column', 'last_row',
        'note_column', 'value_column'
    }),
    'sim_data.run_identity.bank_value_columns': frozenset({'A', 'B'}),
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



CONDITIONAL_KEYS: dict[tuple[str, str], str] = {
    ("sim_data.run_identity.fields[]", "enum"):
        "present exactly when value_type == 'enum', and forbidden otherwise. "
        "Checked by the run-identity validator against the complete locked "
        "record, which is stricter than either 'required' or 'optional'.",
}
"""Keys whose presence is decided by another field, with the rule written down.

A conditional key is NOT an optional one: its presence is required in one case
and refused in the other, and the section validator enforces both directions."""

INTENTIONALLY_OPTIONAL: dict[tuple[str, str], str] = {}
"""Keys that may legitimately be absent, each with a written reason.

DELIBERATELY EMPTY. Every remaining key in this contract is required, INCLUDING
the ones whose canonical value is `null`: `positivity_rule: null` is the
authority SAYING there is no positivity rule, and silent absence does not say
that. Absence and an explicit null are only interchangeable where nothing reads
the distinction, and nothing in this contract is in that position.

An entry here needs a reason, and the deletion sweep fails if this list grows
without one."""

FLEXIBLE_LEAVES: dict[str, str] = {
    "dependence.authority":
        "a prose citation of why no correlation authority exists. The ENFORCED "
        "semantics are the four booleans beside it; this sentence explains them "
        "to a reader and locking its wording would freeze prose, not meaning.",
}
"""Scalar leaves that are descriptive rather than runtime-locked.

Everything else in this contract is settled authority and must not be able to
change while the loader still reports the same accepted Step-1 contract as
valid. The semantic sweep fails if this list grows silently."""


def _check_closed_world(node: Any, path: str, source: Path) -> None:
    """Refuse an unknown key, a mapping at an unknown path, AND a missing key.

    Closure on unknown keys alone was not fail-loud: a deletion sweep found 55
    keys that could simply be removed, so a semantic could be DELETED from the
    authority document and the validator would still report it valid.
    """
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
        excused = {
            key for key in allowed
            if (path, key) in CONDITIONAL_KEYS or (path, key) in INTENTIONALLY_OPTIONAL
        }
        missing = sorted((allowed - excused) - set(node))
        if missing:
            raise SimContractError(
                f"{source}: {path or 'root'} is missing required key(s) {missing}. Absence is "
                "not a value: a key whose canonical content is `null` says there is no such "
                "rule, and deleting it says nothing at all."
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
    _validate_pending_auto_nonce(raw, path)
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
    _validate_sensitivity(raw, path)
    _validate_annual_stochastic(raw, path)
    _validate_selected_confidence_level(raw, path)
    _validate_phase5_bridge(raw, path)

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
    for key, expected in (
        ("rng_version", LOCKED_RNG_VERSION),
        ("sim_method_version", LOCKED_SIM_METHOD_VERSION),
    ):
        value = _req(block, key, f"{path}: versions")
        if not isinstance(value, int) or isinstance(value, bool) or value < 1:
            raise SimContractError(
                f"{path}: versions.{key} must be a positive integer, got {value!r}"
            )
        if value != expected:
            raise SimContractError(
                f"{path}: versions.{key} is {value}; Step 0 settled it at {expected}. A version "
                "bump is an explicit authority change with a written reason, not something a "
                "validator waves through because the new number is also a positive integer."
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
    _require_value(recurrence, "advance", LOCKED_RECURRENCE_ADVANCE, f"{where}: recurrence")

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
    _require_true(life, "prior_successful_publication_untouched", lwhere)
    _validate_allocation_states(life, lwhere)


LOCKED_ALLOCATION_STATES = ("PRE_ALLOCATION", "CONSUMED", "PERSISTENCE_INDETERMINATE")
LOCKED_WRITE_AHEAD_ORDER = (
    "read_current_auto_nonce", "derive_effective_seed",
    "establish_and_verify_pending_auto_nonce", "persist_auto_nonce_plus_one",
    "reconcile_and_clear_pending", "begin_sampling",
)
LOCKED_PENDING_AUTO_NONCE_CELL = "F21"
LOCKED_INDETERMINATE_RESULT = "AUTO_NONCE_INDETERMINATE"


def _is_subsequence(needle: list, haystack: list) -> bool:
    """Every element of `needle`, in order, somewhere in `haystack`."""
    it = iter(haystack)
    return all(item in it for item in needle)


def _validate_allocation_states(life: dict, where: str) -> None:
    """The three-state AUTO allocation authority.

    A `Range.Value2` that RAISES is not proof that Excel wrote nothing, so
    "did the advance persist?" has a third answer. Each state is pinned to the
    exact facts an implementation may claim in it - so that a source which
    calls an unpersisted advance "allocated", or an indeterminate one
    "unconsumed", fails the contract rather than the reviewer.
    """
    _exact_sequence(life.get("allocation_states"), LOCKED_ALLOCATION_STATES,
                    f"{where}: allocation_states")

    semantics = _map(life, "allocation_state_semantics", where)
    swhere = f"{where}: allocation_state_semantics"
    if sorted(semantics) != sorted(LOCKED_ALLOCATION_STATES):
        raise SimContractError(
            f"{swhere}: must describe exactly {sorted(LOCKED_ALLOCATION_STATES)}"
        )

    pre = _map(semantics, "PRE_ALLOCATION", swhere)
    _require_false(pre, "advance_persisted", f"{swhere}: PRE_ALLOCATION")
    _require_false(pre, "nonce_consumed", f"{swhere}: PRE_ALLOCATION")
    _require_false(pre, "sampling_began", f"{swhere}: PRE_ALLOCATION")
    # THE POINT OF THE STATE: a nonce that was never allocated is not "reused".
    _require_true(pre, "retry_may_reuse_the_same_nonce", f"{swhere}: PRE_ALLOCATION")

    used = _map(semantics, "CONSUMED", swhere)
    _require_true(used, "advance_persisted", f"{swhere}: CONSUMED")
    _require_true(used, "nonce_consumed", f"{swhere}: CONSUMED")
    _require_true(used, "sampling_may_begin", f"{swhere}: CONSUMED")
    _require_false(used, "retry_may_reuse_the_same_nonce", f"{swhere}: CONSUMED")

    unknown = _map(semantics, "PERSISTENCE_INDETERMINATE", swhere)
    uwhere = f"{swhere}: PERSISTENCE_INDETERMINATE"
    _require_value(unknown, "advance_persisted", "unknown", uwhere)
    _require_value(unknown, "nonce_consumed", "unknown", uwhere)
    _require_false(unknown, "sampling_began", uwhere)
    _require_true(unknown, "must_not_be_called_allocated", uwhere)
    _require_true(unknown, "must_not_be_called_unconsumed", uwhere)
    # THE TOKEN IS AUDIT ONLY. The durable recovery authority is the sidecar,
    # because the attempt row is rewritten by any later attempt - including a
    # FIXED one that has nothing to do with the pending AUTO transaction.
    _require_value(unknown, "audit_attempt_result", LOCKED_INDETERMINATE_RESULT, uwhere)
    _require_false(unknown, "audit_result_is_the_recovery_authority", uwhere)
    _require_value(unknown, "durable_recovery_authority", "pending_auto_nonce", uwhere)
    _require_true(unknown, "next_run_must_reconcile_before_allocating", uwhere)

    # ONE observation, never a retry loop: a loop would turn an ambiguous COM
    # failure into an unbounded one and still not decide it.
    immediate = _map(life, "immediate_reconciliation", where)
    iwhere = f"{where}: immediate_reconciliation"
    _require_value(immediate, "attempts", 1, iwhere)
    _exact_sequence(immediate.get("after"),
                    ("counter_write_raised", "verification_read_failed",
                     "verification_read_raised"),
                    f"{iwhere}: after")
    # WRITE-AHEAD: the marker is established before the counter is touched.
    _exact_sequence(life.get("write_ahead_order"), LOCKED_WRITE_AHEAD_ORDER,
                    f"{where}: write_ahead_order")
    _require_value(immediate, "observed_m_plus_1", "CONSUMED", iwhere)
    _require_value(immediate, "observed_m", "PRE_ALLOCATION", iwhere)
    _require_value(immediate, "observed_other", "RECOVERY_REQUIRED", iwhere)
    _require_value(immediate, "observation_unavailable", "PERSISTENCE_INDETERMINATE", iwhere)

    # Activated by the durable token ALONE. An ordinary unsuccessful attempt may
    # follow a conclusively persisted advance; conflating them loses exactly the
    # distinction the token exists to keep.
    later = _map(life, "next_run_reconciliation", where)
    nwhere = f"{where}: next_run_reconciliation"
    _require_true(later, "activated_by_pending_auto_nonce_cell", nwhere)
    _require_false(later, "activated_by_attempt_result", nwhere)
    _require_false(later, "activated_by_generic_unsuccessful_result", nwhere)
    _require_false(later, "applies_to_fixed_mode", nwhere)
    _require_value(later, "counter_equals_m_plus_1", "CONSUMED", nwhere)
    _require_value(later, "counter_equals_m", "PRE_ALLOCATION", nwhere)
    _require_value(later, "counter_is_neither", "RECOVERY_REQUIRED", nwhere)
    _require_value(later, "counter_unreadable", "RECOVERY_REQUIRED", nwhere)

    # THE REFINEMENT IS CHECKED, not asserted. `write_ahead_order` must contain
    # every nonce step of the Step-0 `order`, in the same relative order - so
    # the coarse sequence cannot silently become a stale second authority.
    _require_value(life, "write_ahead_order_refines", "order", where)
    dropped = life.get("write_ahead_order_refinement_drops_prefix")
    coarse = list(life.get("order") or ())
    if not coarse or coarse[0] != dropped:
        raise SimContractError(
            f"{where}: write_ahead_order_refinement_drops_prefix must name the "
            f"first step of `order`, which is {coarse[0] if coarse else None!r}"
        )
    if not _is_subsequence(coarse[1:], list(life.get("write_ahead_order") or ())):
        raise SimContractError(
            f"{where}: write_ahead_order {list(life.get('write_ahead_order') or ())} "
            f"does not preserve the remaining steps of `order` {coarse[1:]} in order; "
            "a refinement that drops or reorders a step is a different sequence, "
            "not a refinement"
        )

    # FIXED touches NOTHING in the AUTO transaction. That is what makes a FIXED
    # attempt safe to overwrite the attempt row while a marker stands.
    fixed = _map(life, "fixed_mode", where)
    fwhere = f"{where}: fixed_mode"
    for key in (
        "reads_pending_auto_nonce", "writes_pending_auto_nonce",
        "clears_pending_auto_nonce", "reads_next_auto_nonce",
        "writes_next_auto_nonce", "executes_after_nonce_failpoint",
    ):
        _require_false(fixed, key, fwhere)
    _require_true(fixed, "may_proceed_while_pending_marker_exists", fwhere)
    _require_true(fixed, "may_overwrite_prior_auto_attempt_metadata", fwhere)

    # The fifth token is an AUDIT result. Both unclassified outcomes earn it;
    # recording either as a plain REFUSED would claim the run declined to spend
    # the nonce, which is the one claim the source cannot make.
    token = _map(life, "attempt_result_token", where)
    twhere = f"{where}: attempt_result_token"
    _require_value(token, "token", LOCKED_INDETERMINATE_RESULT, twhere)
    _require_true(token, "audit_only", twhere)
    _require_false(token, "is_the_durable_recovery_authority", twhere)
    _require_true(token, "may_be_overwritten_by_a_later_attempt", twhere)
    # ONE AXIS, ONE MEANING. The token names an unclassifiable transition of THIS
    # attempt; a recovery action is a different question and earns no token.
    _exact_sequence(token.get("recorded_for"), ("PERSISTENCE_INDETERMINATE",),
                    f"{twhere}: recorded_for")
    _require_false(token, "recorded_for_recovery_action", twhere)
    _require_value(token, "other_unsuccessful_outcomes_record", "REFUSED", twhere)
    _require_true(token, "phase5_axis_unchanged", twhere)

    # THE RECOVERY AXIS. An ACTION, never a physical classification - and never
    # a member of allocation_states, which the check below enforces rather than
    # asserts, so a future edit that adds it there is refused.
    action = _map(life, "recovery_action", where)
    awhere = f"{where}: recovery_action"
    _require_value(action, "token", "RECOVERY_REQUIRED", awhere)
    _require_false(action, "is_an_allocation_state", awhere)
    _require_false(action, "derives_physical_consumption", awhere)
    _require_false(action, "may_revise_allocation_state", awhere)
    _require_true(action, "carried_separately_from_allocation_state", awhere)
    _require_false(action, "permits_sampling", awhere)
    if action.get("token") in tuple(life.get("allocation_states") or ()):
        raise SimContractError(
            f"{awhere}: {action.get('token')!r} appears in allocation_states. It is "
            "an action, not a physical classification of a counter transition; "
            "listing it there is what let a cleanup failure erase a proven "
            "CONSUMED observation"
        )
    _exact_sequence(action.get("raised_by"),
                    ("prior_pending_marker_unreadable", "prior_counter_unreadable",
                     "prior_counter_is_neither_m_nor_m_plus_1",
                     "prior_marker_clear_failed", "current_marker_clear_failed",
                     "current_counter_is_neither_m_nor_m_plus_1"),
                    f"{awhere}: raised_by")
    _require_value(action, "allocation_state_when_raised_before_identity_selection",
                   "NOT_APPLICABLE", awhere)
    _require_value(action, "allocation_state_when_raised_by_cleanup", "unchanged",
                   awhere)

    # A raised clear is not proof the marker survived, and no clear failure ever
    # rolls the counter back.
    clear = _map(life, "pending_clear", where)
    cwhere = f"{where}: pending_clear"
    _require_true(clear, "is_a_real_com_write", cwhere)
    _require_false(clear, "raised_clear_proves_marker_remains", cwhere)
    _require_false(clear, "unresolved_cleanup_permits_sampling", cwhere)
    _require_false(clear, "counter_rollback_on_clear_failure", cwhere)

    # The attempt row keeps different facts in different states, and row 27 must
    # never assert physical consumption merely because it holds the number.
    preserves = _map(life, "attempt_metadata_preserves", where)
    pwhere = f"{where}: attempt_metadata_preserves"
    _exact_sequence(preserves.get("known_consumed"),
                    ("consumed_auto_nonce", "effective_seed"),
                    f"{pwhere}: known_consumed")
    _exact_sequence(preserves.get("pre_allocation"),
                    ("attempted_auto_nonce", "effective_seed"),
                    f"{pwhere}: pre_allocation")
    _exact_sequence(preserves.get("persistence_indeterminate"),
                    ("attempted_auto_nonce", "effective_seed",
                     "durable_indeterminate_result"),
                    f"{pwhere}: persistence_indeterminate")

    # THE DECLARED RESIDUAL, stated rather than papered over.
    residual = _map(life, "indeterminate_marker_storage_failure", where)
    rwhere = f"{where}: indeterminate_marker_storage_failure"
    _require_false(residual, "reuse_prevention_depends_on_attempt_row", rwhere)
    _require_true(residual, "reuse_prevention_depends_on_pending_cell", rwhere)
    _require_true(residual, "attempt_row_failure_loses_audit_line_only", rwhere)
    _require_false(residual, "second_write_ahead_log_required", rwhere)


def _validate_pending_auto_nonce(raw: dict, path: Path) -> None:
    """The durable write-ahead recovery marker, and its independence.

    It must not sit anywhere the publication already owns. Column F carries the
    bank-B snapshot, which ends at the last snapshot row; the pending cell is
    the row below it, which is a SHARED counter row and therefore has no bank-B
    twin. That is why the cell is free and why nothing shifts.
    """
    data = _map(raw, "sim_data", path)
    block = _map(data, "pending_auto_nonce", f"{path}: sim_data")
    where = f"{path}: sim_data: pending_auto_nonce"
    _require_value(block, "cell", LOCKED_PENDING_AUTO_NONCE_CELL, where)
    identity = data["run_identity"]
    column = str(block["column"])
    row = int(block["row"])
    if f"{column}{row}" != str(block["cell"]):
        raise SimContractError(f"{where}: cell must equal column & row")
    if column != identity["bank_value_columns"]["B"]:
        raise SimContractError(f"{where}: the sidecar must live in the bank-B column")

    snapshot_rows = [f["row"] for f in identity["fields"] if f.get("group") == "snapshot"]
    if not snapshot_rows:
        # WITHOUT THE SNAPSHOT GROUP THERE IS NOTHING TO PROVE THE SIDECAR CLEAR
        # OF. Falling through would make this validator silently vacuous, which
        # is worse than refusing.
        raise SimContractError(
            f"{where}: no run-identity field is in the 'snapshot' group, so the "
            "bank-B extent cannot be established and the sidecar cannot be "
            "proved free"
        )
    if row <= max(snapshot_rows):
        raise SimContractError(
            f"{where}: {block['cell']} overlaps the bank-B snapshot, which ends at "
            f"row {max(snapshot_rows)}"
        )
    banked = {f["row"] for f in identity["fields"] if f.get("group") == "snapshot"}
    if row in banked:
        raise SimContractError(f"{where}: {block['cell']} is a banked snapshot cell")
    if not identity["fields"]:
        raise SimContractError(f"{where}: the run-identity block declares no fields")
    if row < min(f["row"] for f in identity["fields"]):
        raise SimContractError(f"{where}: the sidecar sits above the identity block")
    records = data["iteration_records"]
    if row >= int(records["header_row"]):
        raise SimContractError(f"{where}: the sidecar collides with the iteration table")

    # THE TWO MEANINGS ARE SETTLED TEXT, not free-form prose. A leaf the
    # validator merely tolerates governs nothing: it would accept a contract
    # saying blank means the opposite of what the source implements, and the
    # closed-world check would still pass because the KEY is known.
    _require_value(block, "label", "Pending AUTO Nonce", where)
    _require_value(
        block, "blank_means",
        "no AUTO allocation transaction requires recovery", where,
    )
    _require_value(
        block, "integer_means",
        "AUTO nonce m is pending; reconcile the counter before allocating", where,
    )
    # It is a machine field, not a claim and not a sentinel.
    _require_value(block, "value_type", "integer_or_blank", where)
    _require_false(block, "is_a_consumed_nonce_claim", where)
    _require_false(block, "is_an_enum_sentinel", where)
    # WRITE-AHEAD, and untouched by FIXED.
    _require_true(block, "written_before_counter_persist", where)
    _require_true(block, "counter_persist_forbidden_until_established", where)
    _require_true(block, "sampling_forbidden_while_unresolved", where)
    _require_true(block, "survives_unrelated_attempts", where)
    _exact_sequence(block.get("cleared_on"),
                    ("counter_equals_m", "counter_equals_m_plus_1"),
                    f"{where}: cleared_on")
    _exact_sequence(block.get("retained_on"),
                    ("counter_is_neither", "counter_unreadable",
                     "observation_unavailable"),
                    f"{where}: retained_on")


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
    _require_value(block, "index_rule", LOCKED_INDEX_RULE, where)
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
    # Verified, not merely shape-checked. Authoritative-looking metadata the
    # validator ignores is worse than no metadata: it invites a reader to trust
    # a binding that does not exist.
    for key, expected in (
        ("a1_p127_sha256", LOCKED_A1_P127_SHA256),
        ("a2_p127_sha256", LOCKED_A2_P127_SHA256),
    ):
        digest = _req_str(block, key, where)
        if not re.fullmatch(r"[0-9a-f]{64}", digest):
            raise SimContractError(f"{where}: {key} must be 64 lowercase hex characters")
        if digest != expected:
            raise SimContractError(
                f"{where}: {key} is {digest}, but the accepted Step-0 hash of the canonical "
                f"matrix text is {expected}. The hash is part of the authority or it should "
                "not be in the contract."
            )


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
    conditions = _map(degenerate, "conditions", dwhere)
    if dict(conditions) != LOCKED_DEGENERATE_CONDITIONS:
        raise SimContractError(
            f"{dwhere}: conditions must be exactly {LOCKED_DEGENERATE_CONDITIONS}, got "
            f"{dict(conditions)}. Detection is FAMILY-SPECIFIC: Uniform's Most Likely is "
            "ignored by accepted Phase-5 D1, so a common `a == m == b` predicate would let an "
            "ignored input decide degeneracy - and therefore RNG consumption, the stream "
            "position, and every later draw on that component."
        )
    _require_false(degenerate, "most_likely_read_by_uniform_degeneracy", dwhere)
    _require_true(degenerate, "detected_before_dispatch", dwhere)
    _require_true(degenerate, "detected_before_parameterisation", dwhere)
    _require_value(degenerate, "returns", "a", dwhere)
    _require_value(degenerate, "uniforms_consumed", 0, dwhere)
    _require_false(degenerate, "sampler_entered", dwhere)
    _require_false(degenerate, "stream_state_changed", dwhere)
    _require_true(degenerate, "applies_to_all_families", dwhere)

    uniform = _map(block, "uniform", where)
    _require_false(uniform, "most_likely_used", f"{where}: uniform")
    _require_false(uniform, "most_likely_affects_degeneracy", f"{where}: uniform")
    _require_false(uniform, "most_likely_affects_uniform_consumption", f"{where}: uniform")
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
    if dict(boundary) != LOCKED_TRIANGULAR_BOUNDARY:
        raise SimContractError(
            f"{twhere}: boundary_cases must be exactly {LOCKED_TRIANGULAR_BOUNDARY}. Which "
            "branch a boundary shape takes is sampling semantics, not documentation."
        )
    _require_true(tri, "rng_endpoints_open", twhere)
    _require_true(tri, "normalised_formulation_required", twhere)
    _require_value(tri, "conditioning_scale", LOCKED_CONDITIONING_SCALE, twhere)

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
    _require_value(pert, "conditioning_scale", LOCKED_CONDITIONING_SCALE, pwhere)
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
    _require_value(bb, "return", LOCKED_BB_RETURN, f"{where}: bb")

    bc = _map(block, "bc", where)
    _require_value(bc, "applies_when", "min(alpha, beta) <= 1", f"{where}: bc")
    _check_orientation(bc, LOCKED_BC_ORIENTATION, f"{where}: bc")
    _exact_sequence(bc.get("literals"), LOCKED_BC_LITERALS, f"{where}: bc.literals")
    _require_value(bc, "acceptance_operator", "greater_than_or_equal", f"{where}: bc")
    # BC was previously validated far more weakly than BB: its expressions and
    # its return rule could be replaced with arbitrary text. BB and BC orient
    # OPPOSITELY, so a free BC return is exactly how a mirrored distribution
    # ships while every other check still passes.
    _exact_sequence(bc.get("per_driver"), LOCKED_BC_PER_DRIVER, f"{where}: bc.per_driver")
    _exact_sequence(bc.get("per_attempt"), LOCKED_BC_PER_ATTEMPT, f"{where}: bc.per_attempt")
    _require_value(bc, "return", LOCKED_BC_RETURN, f"{where}: bc")

    effect = _map(block, "literal_effect", where)
    _require_value(
        effect, "squeeze_literals_affect", "acceptance_decision_only", f"{where}: literal_effect"
    )
    _require_value(
        effect, "logit_form_affects", "returned_sample_value", f"{where}: literal_effect"
    )

    binding = _map(block, "source_binding", where)
    _require_value(
        binding, "evidence_file", LOCKED_CHENG_FORMULATION_EVIDENCE, f"{where}: source_binding"
    )
    _require_value(
        binding, "functions_sha256", LOCKED_CHENG_SOURCE_SHA256, f"{where}: source_binding"
    )
    vectors = _map(block, "conformance_vectors", where)
    _require_value(
        vectors, "evidence_file", LOCKED_CHENG_VECTORS_EVIDENCE, f"{where}: conformance_vectors"
    )
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
    _validate_banks(block, where)
    _validate_transaction(block, where)


def _validate_banks(block: dict, where: str) -> None:
    """Two banks, one active, and a candidate that never touches the active one.

    `prior_successful_publication_survives` was aspirational while a candidate
    overwrote the published rows: a COM failure half way through a million rows
    left a workbook that was neither the old distribution nor the new one, and no
    million-row rollback is a transaction anybody should attempt.
    """
    banks = _map(block, "banks", where)
    bwhere = f"{where}: banks"
    _exact_sequence(banks.get("labels"), LOCKED_BANK_LABELS, f"{bwhere}: labels")
    _require_value(banks, "count", len(LOCKED_BANK_LABELS), bwhere)
    _require_false(banks, "third_bank_permitted", bwhere)
    if banks.get("initial_active_bank") is not None:
        raise SimContractError(
            f"{bwhere}: initial_active_bank must be null. Blank is the ABSENCE of any "
            "successful publication, not a bank; naming one would claim a workbook that "
            "has never run already has a published distribution."
        )
    target = _map(banks, "candidate_target", bwhere)
    if dict(target) != LOCKED_CANDIDATE_TARGET:
        raise SimContractError(
            f"{bwhere}: candidate_target must be exactly {LOCKED_CANDIDATE_TARGET}, got "
            f"{dict(target)}. The first success targets A, and every success afterwards "
            "targets whichever bank is NOT active."
        )
    _require_false(banks, "candidate_writes_to_active_bank", bwhere)
    _require_false(banks, "inactive_bank_is_published", bwhere)
    _require_true(banks, "inactive_bank_is_staging_storage", bwhere)
    _require_false(banks, "temporary_worksheet_required", bwhere)
    _require_false(banks, "duplicate_workbook_required", bwhere)
    _require_true(banks, "row_axis_shared_by_both_banks", bwhere)


def _validate_transaction(block: dict, where: str) -> None:
    transaction = _map(block, "transaction", where)
    twhere = f"{where}: transaction"
    _exact_sequence(transaction.get("order"), LOCKED_TRANSACTION_ORDER, f"{twhere}: order")
    order = tuple(transaction.get("order") or ())
    if order[-1] != "final_commit_shared_block_including_active_bank":
        raise SimContractError(
            f"{twhere}: the active-bank switch must be the LAST step. A switch before the "
            "candidate bank is verified publishes an unverified distribution."
        )
    if order.index("verify_inactive_bank_against_staged_package") >= order.index(
            "final_commit_shared_block_including_active_bank"):
        raise SimContractError(f"{twhere}: verification must precede the final commit")
    if order.index("validate_pre_allocation_prerequisites") >= order.index(
            "allocate_auto_nonce_when_auto"):
        raise SimContractError(
            f"{twhere}: every pre-allocation prerequisite must be checked BEFORE an AUTO "
            "nonce is consumed. A run that can never be committed must not burn a "
            "sequence."
        )
    _require_value(transaction, "final_commit_range", LOCKED_FINAL_COMMIT_RANGE, twhere)
    _require_true(transaction, "final_commit_is_one_write", twhere)
    _exact_sequence(transaction.get("final_commit_fields"), LOCKED_FINAL_COMMIT_FIELDS,
                    f"{twhere}: final_commit_fields")
    if tuple(transaction.get("final_commit_fields") or ())[-1] != "active_bank":
        raise SimContractError(
            f"{twhere}: active_bank must be the last field of the final commit block"
        )
    _require_true(transaction, "prior_final_commit_block_captured_before_write", twhere)
    _require_true(transaction, "final_commit_failure_restores_prior_block", twhere)
    _require_false(transaction, "million_row_rollback_required", twhere)
    _require_false(transaction, "results_is_a_written_transaction", twhere)

    allocation = _map(block, "run_id_allocation", where)
    awhere = f"{where}: run_id_allocation"
    _require_value(allocation, "candidate_value", "last_run_id + 1", awhere)
    _require_true(allocation, "held_locally_until_commit", awhere)
    _require_value(allocation, "allocated_by", "successful_final_commit", awhere)
    _require_true(allocation, "headroom_checked_before_auto_allocation", awhere)

    failure = _map(block, "failure_semantics", where)
    fwhere = f"{where}: failure_semantics"
    before = _map(failure, "refusal_before_auto_allocation", fwhere)
    _require_false(before, "next_auto_nonce_advanced", f"{fwhere}: refusal_before_auto_allocation")
    _require_false(before, "active_bank_changed", f"{fwhere}: refusal_before_auto_allocation")
    _require_false(before, "successful_banks_changed", f"{fwhere}: refusal_before_auto_allocation")
    _require_true(before, "attempt_metadata_updated", f"{fwhere}: refusal_before_auto_allocation")
    after = _map(failure, "refusal_or_failure_after_auto_allocation", fwhere)
    awh = f"{fwhere}: refusal_or_failure_after_auto_allocation"
    # NOT rolled back, and deliberately: a consumed AUTO sequence is consumed.
    _require_true(after, "next_auto_nonce_advanced", awh)
    _require_false(after, "active_bank_changed", awh)
    _require_false(after, "successful_banks_changed", awh)
    _require_true(after, "attempt_metadata_updated", awh)
    inactive = _map(failure, "inactive_bank_write_failure", fwhere)
    iwh = f"{fwhere}: inactive_bank_write_failure"
    _require_false(inactive, "active_bank_changed", iwh)
    _require_true(inactive, "prior_publication_remains_authoritative", iwh)
    _require_false(inactive, "corrupted_candidate_has_semantic_standing", iwh)
    commit = _map(failure, "final_commit_failure", fwhere)
    cwh = f"{fwhere}: final_commit_failure"
    _require_true(commit, "prior_block_restored", cwh)
    _require_false(commit, "active_bank_changed", cwh)


def _validate_command_surface(raw: dict, path: Path) -> None:
    block = _map(raw, "command_surface", path)
    where = f"{path}: command_surface"
    _require_value(block, "automation_endpoint", "PCCM_RunSimulation", where)
    for flag in (
        "user_facing_run_button_in_phase_6",
        "msgbox_introduced_by_phase_6",
        "userform_introduced_by_phase_6",
        "ribbon_introduced_by_phase_6",
    ):
        _require_false(block, flag, where)
    # SETTLED at Step 11A, before the module exists, so it cannot invent a name.
    _require_true(block, "read_accessor_names_settled", where)
    _exact_sequence(block.get("read_accessors"), LOCKED_READ_ACCESSORS,
                    f"{where}: read_accessors")
    semantics = _map(block, "read_accessor_semantics", where)
    if dict(semantics) != LOCKED_READ_ACCESSOR_SEMANTICS:
        raise SimContractError(
            f"{where}: read_accessor_semantics must be exactly the accepted wording, in "
            f"order, got {dict(semantics)}. These are the authority a later implementation "
            "reads, not commentary: 'the stored fingerprint' and 'the recomputed current "
            "fingerprint' are different procedures and must stay different sentences."
        )
    if tuple(semantics) != LOCKED_READ_ACCESSORS:
        raise SimContractError(
            f"{where}: read_accessor_semantics must describe the accessors in the settled "
            f"order, got {list(semantics)}"
        )
    endpoint = _req_str(block, "automation_endpoint", where)
    if endpoint in LOCKED_READ_ACCESSORS:
        raise SimContractError(f"{where}: the run endpoint may not also be a read accessor")
    _require_false(block, "run_id_public_accessor_required_in_phase_6", where)
    _require_false(block, "effective_seed_public_accessor_required_in_phase_6", where)
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
    _validate_request_sim_grammar(sim, f"{where}: sim_section")
    _require_true(block, "auto_blank_seed_remains_recomputable", where)


def _validate_request_sim_grammar(sim: dict, where: str) -> None:
    """The SIM extension token by token.

    Before this closure the contract locked the SEMANTIC fields and their order
    and stopped there, so several byte-distinct streams satisfied it: F_I versus
    F_N for iterations, one record versus five, an AUTO seed omitted versus
    blank versus zero, versions as integers versus text. A later implementation
    would have picked one by accident and that accident would have become the
    identity of every stored request fingerprint.
    """
    _require_value(sim, "record_count", LOCKED_REQUEST_RECORD_COUNT, where)
    _require_false(sim, "encoded_field_names", where)

    types = _map(sim, "field_types", where)
    if dict(types) != LOCKED_REQUEST_FIELD_TYPES:
        raise SimContractError(
            f"{where}: field_types must be exactly {LOCKED_REQUEST_FIELD_TYPES}, got "
            f"{dict(types)}. Every integer identity here is F_I: a count, a seed and a "
            "version are structural facts, and F_N would let a version of 1 encode "
            "identically to a Double of 1."
        )
    if set(types) != set(LOCKED_SIM_FIELDS):
        raise SimContractError(
            f"{where}: field_types must type exactly the declared fields"
        )

    effective = _map(sim, "effective_records", where)
    if tuple(effective) != tuple(LOCKED_SEED_MODES):
        raise SimContractError(
            f"{where}: effective_records must describe exactly {list(LOCKED_SEED_MODES)}, "
            f"in that order, got {list(effective)}"
        )
    for mode, expected in LOCKED_REQUEST_EFFECTIVE.items():
        shape = _map(effective, mode, f"{where}: effective_records")
        ewhere = f"{where}: effective_records.{mode}"
        _exact_sequence(shape.get("fields"), expected, f"{ewhere}: fields")
        _require_value(shape, "field_count", len(expected), ewhere)
        unknown = [name for name in expected if name not in LOCKED_REQUEST_FIELD_TYPES]
        if unknown:  # pragma: no cover - _exact_sequence already pins the list
            raise SimContractError(f"{ewhere}: untyped field(s) {unknown}")
    auto_fields = tuple(effective["AUTO"].get("fields") or ())
    if "supplied_seed" in auto_fields:
        raise SimContractError(
            f"{where}: effective_records.AUTO carries supplied_seed. AUTO means the field "
            "DOES NOT EXIST - not F_I(0), not F_S(\"\"), not null and not the previous "
            "effective seed. That absence is why two AUTO runs of the same question share "
            "one request fingerprint and stay CURRENT."
        )
    if "supplied_seed" not in tuple(effective["FIXED"].get("fields") or ()):
        raise SimContractError(
            f"{where}: effective_records.FIXED omits supplied_seed, which is the only thing "
            "distinguishing one FIXED request from another"
        )

    grammar = _map(sim, "grammar", where)
    if dict(grammar) != LOCKED_REQUEST_GRAMMAR:
        raise SimContractError(
            f"{where}: grammar must be exactly {LOCKED_REQUEST_GRAMMAR}. The request "
            "fingerprint is now locked token by token, to the same standard result_digest "
            "has always carried."
        )
    for production in grammar.values():
        for banned in ("PCCM-FP", "FP_VERSION", "SIM_FP_VERSION", "REQUEST_FP_VERSION"):
            if banned in str(production):
                raise SimContractError(
                    f"{where}: the grammar repeats {banned!r} inside the SIM extension. The "
                    "extension is a SECTION of the accepted PCCM-FP stream; it carries no "
                    "stream tag and no stream version of its own."
                )
        for excluded in LOCKED_SIM_EXCLUDED:
            if excluded in str(production):
                raise SimContractError(
                    f"{where}: the grammar encodes the excluded field {excluded!r}"
                )

    _require_value(sim, "auto_supplied_seed_representation", LOCKED_REQUEST_SEED_ABSENT, where)
    _require_false(sim, "stream_tag_repeated_in_extension", where)
    _require_false(sim, "stream_version_repeated_in_extension", where)
    _require_value(sim, "stream_tag_owner", LOCKED_REQUEST_STREAM_TAG_OWNER, where)
    _require_value(
        sim, "supplied_seed_domain_owner", LOCKED_REQUEST_SEED_DOMAIN_OWNER, where
    )


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
    if dict(grammar) != LOCKED_DIGEST_GRAMMAR:
        raise SimContractError(
            f"{where}: grammar must be exactly {LOCKED_DIGEST_GRAMMAR}. D6-17 locks the stream "
            "token by token; a grammar the validator does not check is a grammar the "
            "implementation can choose."
        )
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
    _exact_sequence(block.get("bank"), LOCKED_BANK_LABELS, f"{where}: bank")
    banks = tuple(raw["publication"]["banks"].get("labels") or ())
    if banks != tuple(block.get("bank") or ()):
        raise SimContractError(
            f"{where}: bank labels {list(block.get('bank') or ())} disagree with "
            f"publication.banks.labels {list(banks)}"
        )


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
    if dict(definitions) != LOCKED_SIM_STATE_DEFINITIONS:
        raise SimContractError(
            f"{where}: the state definitions must be exactly the accepted corrected wording. "
            "They are the authority a later implementation reads, not commentary."
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


def _validate_summary_statistics(block: dict, where: str, header_row: int) -> None:
    """The Step-9 summary, PERSISTED.

    `Results derives from _SimData` and `modSimStats` owns these numbers. Both
    are true together only if the summary is stored: a worksheet AVERAGE or
    PERCENTILE would be a second statistics engine, and reading the iteration
    rows back to recompute would make the presentation layer a calculator.
    """
    summary = _map(block, "summary_statistics", where)
    swhere = f"{where}: summary_statistics"
    _require_value(summary, "label_column", LOCKED_SUMMARY_COLUMNS["label_column"], swhere)
    columns = _map(summary, "bank_value_columns", swhere)
    for bank in LOCKED_BANK_LABELS:
        if dict(_map(columns, bank, swhere)) != LOCKED_SUMMARY_COLUMNS[bank]:
            raise SimContractError(
                f"{swhere}: bank {bank} columns must be {LOCKED_SUMMARY_COLUMNS[bank]}"
            )
    _require_value(summary, "source", "modSimStats", swhere)
    _require_false(summary, "recomputed_from_worksheet_data", swhere)
    metrics = _seq(summary, "metrics", swhere)
    actual = tuple(
        (_req_str(m, "key", swhere), _req_int(m, "row", swhere),
         m.get("label"), _req_str(m, "source", swhere))
        for m in metrics
    )
    for key, _row, label, _source in actual:
        if key.startswith("quantile_"):
            if label is not None:
                raise SimContractError(
                    f"{swhere}: rung {key} carries the label {label!r}. The selectable "
                    "ladder belongs to input_contract.yaml; a label spelled here is a "
                    "second ladder that can drift from its owner."
                )
        elif not isinstance(label, str) or not label.strip():
            raise SimContractError(f"{swhere}: {key} carries no label")
    if actual != LOCKED_SUMMARY_METRICS:
        raise SimContractError(
            f"{swhere}: the persisted summary must be exactly the accepted Step-9 surface - "
            f"mean, sample deviation, minimum, the eleven projected rungs, maximum and the "
            f"deterministic base - in order, got {[list(a) for a in actual]}"
        )
    _require_value(summary, "first_row", actual[0][1], swhere)
    _require_value(summary, "last_row", actual[-1][1], swhere)
    if [a[1] for a in actual] != list(range(actual[0][1], actual[-1][1] + 1)):
        raise SimContractError(f"{swhere}: the metric rows must be contiguous and in order")
    if actual[-1][1] >= header_row:
        raise SimContractError(
            f"{swhere}: the summary block reaches row {actual[-1][1]}, at or below the "
            f"iteration header row {header_row}"
        )


def _validate_contingency_ladder(block: dict, where: str, header_row: int) -> None:
    """THE WHOLE LADDER, precomputed before the candidate bank may commit.

    Selected Confidence Level is reporting-only and may move without a rerun. A
    publication holding only the selected rung would force either a rerun or a
    worksheet subtraction the moment it did.
    """
    ladder = _map(block, "contingency_ladder", where)
    cwhere = f"{where}: contingency_ladder"
    _require_value(ladder, "label_column", LOCKED_CONTINGENCY_COLUMNS["label_column"], cwhere)
    columns = _map(ladder, "bank_value_columns", cwhere)
    for bank in LOCKED_BANK_LABELS:
        if dict(_map(columns, bank, cwhere)) != LOCKED_CONTINGENCY_COLUMNS[bank]:
            raise SimContractError(
                f"{cwhere}: bank {bank} columns must be {LOCKED_CONTINGENCY_COLUMNS[bank]}"
            )
    _require_value(ladder, "source", "SimStatsContingency", cwhere)
    _require_value(ladder, "baseline", LOCKED_CONTINGENCY_BASELINE, cwhere)
    _require_false(ladder, "worksheet_subtraction_permitted", cwhere)
    _require_true(ladder, "computed_for_whole_ladder_before_commit", cwhere)
    _require_true(ladder, "all_values_representable_required_before_commit", cwhere)
    _require_true(ladder, "fixed_rung_persisted_though_not_selectable", cwhere)
    rungs = _seq(ladder, "rungs", cwhere)
    actual = tuple(
        (_req_str(r, "key", cwhere), _req_int(r, "row", cwhere), r.get("label"))
        for r in rungs
    )
    for key, _row, label in actual:
        if label is not None:
            raise SimContractError(
                f"{cwhere}: rung {key} spells its label; the ladder belongs to "
                "input_contract.yaml"
            )
    expected = tuple(
        (key, 8 + offset, label)
        for offset, (key, _row, label, _source) in enumerate(
            m for m in LOCKED_SUMMARY_METRICS if m[0].startswith("quantile_"))
    )
    if any(row[2] is not None for row in expected):  # pragma: no cover - locked above
        raise SimContractError(f"{cwhere}: a rung label is spelled in this contract")
    if actual != expected:
        raise SimContractError(
            f"{cwhere}: the contingency ladder must carry EVERY reported rung, in the "
            f"projected order, got {[list(a) for a in actual]}. Storing only the selected "
            "rung would force a rerun or a worksheet subtraction the moment the selector "
            "moved."
        )
    _require_value(ladder, "first_row", actual[0][1], cwhere)
    _require_value(ladder, "last_row", actual[-1][1], cwhere)
    if actual[-1][1] >= header_row:
        raise SimContractError(f"{cwhere}: the ladder reaches the iteration header row")


def _validate_selected_confidence_level(raw: dict, path: Path) -> None:
    block = _map(raw, "selected_confidence_level", path)
    where = f"{path}: selected_confidence_level"
    _require_value(block, "source", "inpSelectedConfidenceLevel", where)
    for flag in ("participates_in_request_fingerprint", "participates_in_execution_validity",
                 "participates_in_auto_allocation", "participates_in_state_derivation",
                 "change_requires_rerun", "invalid_selector_invalidates_simulation",
                 "unselected_state_introduced"):
        _require_false(block, flag, where)
    _require_true(block, "invalid_selector_blanks_selected_reporting_rows", where)


def _validate_phase5_bridge(raw: dict, path: Path) -> None:
    """ONE reusable surface into the accepted Phase-5 preparation.

    A second construction of DriverFactors, the analytical fingerprint, the
    deterministic base or the applied timeline is a second answer to the same
    question, and the two would drift silently.
    """
    block = _map(raw, "phase5_bridge", path)
    where = f"{path}: phase5_bridge"
    _require_value(block, "owner_module", "modCalcReport", where)
    procedure = _req_str(block, "procedure", where)
    if procedure != "CalcPrepareSimulationInputs":
        raise SimContractError(
            f"{where}: procedure must be 'CalcPrepareSimulationInputs', got {procedure!r}"
        )
    if procedure.startswith("PCCM_"):
        raise SimContractError(f"{where}: the bridge is internal, not an endpoint")
    _require_false(block, "is_automation_endpoint", where)
    _require_false(block, "name_prefix_pccm", where)
    _require_value(block, "reuses_private_preparation", "PrepareCurrentCalculation", where)
    _require_value(block, "requires_phase5_status", "CURRENT", where)
    _require_false(block, "writes_to_calc_sheet", where)
    _require_false(block, "updates_phase5_status_or_attempt_metadata", where)
    _require_false(block, "duplicates_factor_mathematics", where)
    _require_true(block, "zero_driver_model_succeeds", where)
    _exact_sequence(block.get("returns"), LOCKED_PHASE5_BRIDGE_RETURNS, f"{where}: returns")
    _require_true(block, "analytical_fingerprint_is_current_not_stored", where)


# ---------------------------------------------------------------------------
# PHASE 7 - SENSITIVITY.
#
# Every decision the Phase-7 authority settled is asserted here. A section whose
# keys are merely spelled correctly governs nothing, and this file's own rule is
# that a field which looks authoritative while enforcing nothing is worse than
# no field at all.
# ---------------------------------------------------------------------------
def _validate_sensitivity(raw: dict, path: Path) -> None:
    where = f"{path}: sensitivity"
    block = _map(raw, "sensitivity", where)
    _require_value(block, "phase", 7, where)
    _require_value(block, "kind", "observational_post_processing", where)
    _require_false(block, "independent_monte_carlo_permitted", where)
    _require_value(block, "basis_run",
                   "the accepted successful snapshot in the active bank", where)

    dwhere = f"{where}: drivers"
    drivers = _map(block, "drivers", dwhere)
    _require_true(drivers, "one_per_cost_line", dwhere)
    _require_true(drivers, "one_per_risk", dwhere)
    _require_false(drivers, "category_aggregation", dwhere)
    _require_value(drivers, "identity", "permanent_id", dwhere)
    _require_false(drivers, "identity_is_worksheet_row", dwhere)

    cwhere = f"{where}: contribution"
    contribution = _map(block, "contribution", cwhere)
    _require_value(contribution, "measure", "nominal", cwhere)
    # THE EXPRESSIONS ARE CHECKED AGAINST THEIR OWNER, not against a literal
    # repeated here. If `contribution` ever changes, this must move with it or
    # fail - which is the whole point of naming one owner.
    owner = _map(raw, "contribution", f"{path}: contribution")
    cost_owner = _map(owner, "cost_line", f"{path}: contribution.cost_line")
    risk_owner = _map(owner, "risk", f"{path}: contribution.risk")
    _require_value(contribution, "cost_line", cost_owner.get("nominal"), cwhere)
    expected_risk = (f"occurred ? {risk_owner.get('nominal_when_occurred')} : "
                     f"{risk_owner.get('nominal_when_not_occurred')}")
    _require_value(contribution, "risk", expected_risk, cwhere)
    # THE ONE OWNER. The engine's `contribution` section already states this
    # arithmetic; a second copy is the duplicate source of truth the whole
    # results architecture forbids.
    _require_value(contribution, "expression_owner", "contribution", cwhere)
    _require_false(contribution, "reimplementation_permitted", cwhere)
    _require_true(contribution, "risk_occurrence_and_severity_are_one_driver", cwhere)
    _require_false(contribution, "correlation_against_raw_severity_permitted", cwhere)

    rwhere = f"{where}: replay"
    replay = _map(block, "replay", rwhere)
    _require_value(replay, "granularity", "driver", rwhere)
    _require_value(replay, "resets_from", "component_initial_state", rwhere)
    # The withdrawn seek claim, kept withdrawn: a rejection sampler consumes a
    # variable number of uniforms, so iteration j is reached by advancing to it.
    _require_false(replay, "random_access_seek", rwhere)
    _require_true(replay, "sequential_advance_required", rwhere)
    _require_value(replay, "cost_line_streams", 1, rwhere)
    _require_value(replay, "risk_streams", 2, rwhere)
    _require_true(replay, "risk_streams_paired_by_iteration", rwhere)
    _require_false(replay, "unrelated_drivers_advanced", rwhere)
    _require_false(replay, "retains_driver_matrix", rwhere)
    _require_value(replay, "concurrent_driver_columns_retained", 1, rwhere)

    swhere = f"{where}: statistic"
    statistic = _map(block, "statistic", swhere)
    _require_value(statistic, "name", "spearman_rank_correlation", swhere)
    _require_value(statistic, "definition",
                   "pearson(midrank(driver_contribution), midrank(total_nominal))", swhere)
    _require_value(statistic, "tie_rule", "average_ranks", swhere)
    # THE SHORTCUT IS WRONG, NOT MERELY SLOWER. A Risk at p = 0.2 puts roughly
    # 80% of its column on one tied value.
    _require_false(statistic, "no_ties_shortcut_permitted", swhere)
    _require_true(statistic, "total_ranks_computed_once", swhere)
    _require_true(statistic, "total_ranks_reused_across_drivers", swhere)
    _require_value(statistic, "sorting", "on_copies_only", swhere)
    _require_false(statistic, "source_arrays_mutated", swhere)
    _require_true(statistic, "iteration_correspondence_preserved", swhere)

    zwhere = f"{where}: zero_variance"
    zero = _map(block, "zero_variance", zwhere)
    _require_value(zero, "status_label", "n/a - no variance", zwhere)
    _require_false(zero, "rho_reported", zwhere)
    # "No association was measured" and "no measurement was possible" are
    # different facts, and rho = 0 asserts the first one.
    _require_false(zero, "reported_as_zero_rho", zwhere)
    _require_true(zero, "excluded_from_ranking", zwhere)
    _require_true(zero, "excluded_from_tornado_input", zwhere)
    _require_true(zero, "retained_diagnostically", zwhere)

    kwhere = f"{where}: ranking"
    ranking = _map(block, "ranking", kwhere)
    _require_value(ranking, "order_by", "absolute_rho", kwhere)
    _require_value(ranking, "direction", "descending", kwhere)
    _require_true(ranking, "signed_rho_retained", kwhere)
    _exact_sequence(ranking.get("direction_labels"), ("+", "-"), f"{kwhere}: direction_labels")
    # Top-N is a CHART decision, and charts are Phase 8. Truncating here would
    # discard data the later phase is entitled to choose from.
    _require_false(ranking, "top_n_truncation", kwhere)
    _require_value(ranking, "population", "every eligible non_zero_variance driver", kwhere)
    # P7-2 DELTA. Without a fixed tie-break, equal |rho| leaves the order to the
    # sort, and the same model yields two different tables.
    _require_value(ranking, "tie_break", "permanent_id", kwhere)
    _require_value(ranking, "tie_break_direction", "ascending", kwhere)
    _require_value(ranking, "tie_break_comparison", LOCKED_ID_COMPARISON, kwhere)
    _require_false(ranking, "tie_break_uses_worksheet_row", kwhere)
    _require_false(ranking, "tie_break_uses_supply_order", kwhere)

    pwhere = f"{where}: sampling"
    sampling = _map(block, "sampling", pwhere)
    _require_true(sampling, "uses_all_iterations", pwhere)
    _require_false(sampling, "subsampling_contracted", pwhere)
    for key in ("iteration_cap", "sensitivity_sample_size"):
        if sampling.get(key) is not None:
            raise SimContractError(
                f"{pwhere}: {key} is set. No iteration cap or sample size is authorised, "
                "and under replay the dominant cost is stream advancement, which a "
                "subsample does not avoid. Measured evidence changes this contract first."
            )
    _require_false(sampling, "unmeasured_performance_safeguard_permitted", pwhere)
    iwhere = f"{pwhere}: index_set_rule_if_ever_adopted"
    index_rule = _map(sampling, "index_set_rule_if_ever_adopted", iwhere)
    _require_true(index_rule, "shared_index_set", iwhere)
    _require_true(index_rule, "same_indices", iwhere)
    _require_true(index_rule, "same_order", iwhere)
    _require_false(index_rule, "independently_selected_samples_permitted", iwhere)
    _require_false(index_rule, "full_run_statistics_recomputed_from_subset", iwhere)

    nwhere = f"{where}: interpretation"
    interpretation = _map(block, "interpretation", nwhere)
    _require_value(interpretation, "measures", "monotone_association", nwhere)
    _require_false(interpretation, "measures_variance_contribution", nwhere)
    # rho^2 is not a share of variance. Presenting it as a percentage would be a
    # quantitative claim Spearman does not support.
    _require_false(interpretation, "rho_squared_as_variance_share_permitted", nwhere)
    _require_false(interpretation, "percentage_contribution_permitted", nwhere)
    _require_value(interpretation, "inter_driver_correlation_owner", "dependence", nwhere)

    bwhere = f"{where}: identity_binding"
    binding = _map(block, "identity_binding", bwhere)
    for field in ("run_id", "effective_seed", "request_fingerprint", "result_digest"):
        if field not in (binding.get("stamped_with") or ()):
            raise SimContractError(
                f"{bwhere}: stamped_with omits {field!r}. Sensitivity that cannot be tied "
                "to the run it describes can be shown against a different one."
            )
    _require_value(binding, "stored_in", "the same bank as the run it describes", bwhere)
    _require_true(binding, "valid_only_for_the_stamped_run", bwhere)

    twhere = f"{where}: state_safety"
    safety = _map(block, "state_safety", twhere)
    for key in ("consumes_run_id", "advances_auto_nonce",
                "touches_pending_auto_nonce_marker", "writes_attempt_row",
                "mutates_successful_snapshot", "rewrites_iteration_records",
                "changes_result_digest"):
        _require_false(safety, key, twhere)

    vwhere = f"{where}: display"
    display = _map(block, "display", vwhere)
    _require_value(display, "state_owner", "sim_state", vwhere)
    _require_value(
        display, "presented_as_current_when",
        "sim_state is CURRENT and the stored sensitivity stamp equals the active bank "
        "successful stamp", vwhere)
    _require_false(display, "presented_as_current_when_stale", vwhere)
    _require_false(display, "presented_as_current_when_invalid", vwhere)
    _require_true(display, "stale_must_be_labelled", vwhere)
    # Inherited from sim_state.on_failure.prior_sim_data_preserved.
    _require_false(display, "refused_attempt_destroys_prior_sensitivity", vwhere)
    _require_true(display, "prior_sensitivity_preserved_on_failure", vwhere)


# ---------------------------------------------------------------------------
# PHASE 7 - ANNUAL STOCHASTIC OUTPUT.
# ---------------------------------------------------------------------------
def _validate_annual_stochastic(raw: dict, path: Path) -> None:
    where = f"{path}: annual_stochastic"
    block = _map(raw, "annual_stochastic", where)
    _require_value(block, "phase", 7, where)
    _require_true(block, "contracted", where)

    fwhere = f"{where}: per_year_factor"
    factor = _map(block, "per_year_factor", fwhere)
    # The per-year factor is a DECOMPOSITION of an accepted published number and
    # sums back to it. Calling it a new input would put a second owner on Knom.
    _require_false(factor, "is_a_new_input", fwhere)
    _require_value(factor, "nominal", "Knom_y = FX * w_y * infl_y", fwhere)
    _require_value(factor, "pv", "Kpv_y = FX * w_y * infl_y * disc_y", fwhere)
    _require_value(factor, "factor_owner", "calc_contract.yaml", fwhere)
    _require_value(factor, "reconciles", "sum_y Knom_y = Knom", fwhere)
    _exact_sequence(factor.get("decomposition_of"), ("Knom", "Kpv"),
                    f"{fwhere}: decomposition_of")

    vwhere = f"{where}: iteration_annual_vector"
    vector = _map(block, "iteration_annual_vector", vwhere)
    _require_value(vector, "nominal",
                   "A_j(y) = sum_d sample_d_j * Quantity_d * Knom_d_y", vwhere)
    _require_value(vector, "pv",
                   "A_j(y) = sum_d sample_d_j * Quantity_d * Kpv_d_y", vwhere)
    _require_value(vector, "identity", "sum_y A_j(y) = iteration total for j", vwhere)
    _require_value(vector, "identity_owner", "docs/phase5_plan.md I3c", vwhere)

    rwhere = f"{where}: retention"
    retention = _map(block, "retention", rwhere)
    # The same trap `per_driver_samples` refuses, on the other axis.
    _require_false(retention, "persisted_iteration_by_year_matrix", rwhere)
    _require_false(retention, "retained_in_memory_matrix", rwhere)
    _require_value(retention, "strategy", "block_replay", rwhere)
    _require_value(retention, "block_axis", "project_year", rwhere)
    _require_true(retention, "block_width_configurable", rwhere)
    _require_value(retention, "passes", "ceil(applied_duration / block_width)", rwhere)
    _exact_sequence(retention.get("retained"),
                    ("annual_percentile_ladder", "selected_px_annual_profile"),
                    f"{rwhere}: retained")

    dwhere = f"{where}: annual_distributions"
    distributions = _map(block, "annual_distributions", dwhere)
    _require_value(distributions, "quantile_method_owner", "statistics.percentile", dwhere)
    _require_value(distributions, "method", LOCKED_PERCENTILE_METHOD, dwhere)
    _require_value(distributions, "sorting", "on_copies_only", dwhere)
    _require_true(distributions, "per_year", dwhere)
    # (a) IS NOT (b). Per-year percentiles do not sum to the reported total
    # percentile, so presenting them as a profile would be a reconciliation
    # error. The two are named separately so that cannot happen quietly.
    _require_false(distributions, "sums_to_total_percentile", dwhere)
    _require_false(distributions, "is_a_selected_px_profile", dwhere)
    _require_value(distributions, "ladder_owner", "statistics", dwhere)
    _exact_sequence(distributions.get("measures"), ("nominal", "pv"),
                    f"{dwhere}: measures")

    pwhere = f"{where}: selected_px_profile"
    profile = _map(block, "selected_px_profile", pwhere)
    _require_value(profile, "definition", "convex_type_7_blend", pwhere)
    _require_value(
        profile, "formula",
        "Profile_Px(y) = (1 - f) * AnnualVector_lo(y) + f * AnnualVector_hi(y)", pwhere)
    _require_value(profile, "position_owner", "statistics.percentile", pwhere)
    _require_true(profile, "degenerates_to_single_iteration_when_f_is_zero", pwhere)
    _require_value(profile, "reconciliation_identity",
                   "sum_y Profile_Px(y) = reported Px", pwhere)
    _require_false(profile, "nearest_rank_permitted", pwhere)
    _require_false(profile, "per_year_percentile_as_profile_permitted", pwhere)
    _require_false(profile, "other_definitions_permitted", pwhere)
    _require_value(profile, "lo_hi_f_source",
                   "the same type-7 position used for the reported total Px", pwhere)
    _require_value(profile, "reconciliation_rule_owner",
                   "docs/phase5_plan.md section 15", pwhere)
    _exact_sequence(profile.get("measures"), ("nominal", "pv"), f"{pwhere}: measures")

    hwhere = f"{where}: phase_8_handoff"
    handoff = _map(block, "phase_8_handoff", hwhere)
    _require_false(handoff, "presentation_in_phase_7", hwhere)
    _exact_sequence(handoff.get("provides"),
                    ("annual_percentile_ladder", "selected_px_annual_profile"),
                    f"{hwhere}: provides")
    # Every one of these is Phase 8, and saying so here is what keeps Phase 7
    # from drifting into presentation.
    for key in ("annual_cash_flow_presentation_owner", "dashboard_owner", "chart_owner"):
        _require_value(handoff, key, "phase 8", hwhere)


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
    # PHASE 7. Annual simulated samples are now contracted - by
    # `annual_stochastic`, not by Results. The `deferred` list above is what
    # keeps Results itself presentation-free, and it is unchanged.
    _require_true(block, "annual_simulated_samples_contracted", where)

    # RESULTS IS PRESENTATION, and this is where that stops being a wish.
    presentation = _map(block, "presentation", where)
    pwhere = f"{where}: presentation"
    _require_false(presentation, "written_by_the_run", pwhere)
    _require_true(presentation, "materialised_by_stage_a", pwhere)
    _exact_sequence(presentation.get("reads_only"),
                    ("the active _SimData bank", "inpSelectedConfidenceLevel"),
                    f"{pwhere}: reads_only")
    for flag in ("computes_statistics", "computes_contingency", "recomputes_quantiles",
                 "contingency_by_subtraction_on_results", "reads_a_fixed_bank"):
        _require_false(presentation, flag, pwhere)
    _exact_sequence(presentation.get("forbidden_functions"),
                    LOCKED_RESULTS_FORBIDDEN_FUNCTIONS,
                    f"{pwhere}: forbidden_functions")
    _exact_sequence(presentation.get("run_stamp_fields"), LOCKED_RESULTS_RUN_STAMP_FIELDS,
                    f"{pwhere}: run_stamp_fields")
    # CROSS-VALIDATION: every field Results presents must be a real `_SimData`
    # field identity, and every one of them must come from the right group.
    identities = {row[0]: row[2] for row in LOCKED_RUN_IDENTITY}
    for key in LOCKED_RESULTS_RUN_STAMP_FIELDS:
        if key not in identities:
            raise SimContractError(
                f"{pwhere}: run_stamp_fields names {key!r}, which is not a _SimData field"
            )
        if identities[key] not in ("snapshot", "derived"):
            raise SimContractError(
                f"{pwhere}: {key!r} is a {identities[key]} field; the Run Stamp presents "
                "the successful snapshot and the derived status, not the counters"
            )
    snapshot_keys = {k for k, group in identities.items() if group == "snapshot"}
    missing = sorted(snapshot_keys - set(LOCKED_RESULTS_RUN_STAMP_FIELDS))
    if missing:
        raise SimContractError(
            f"{pwhere}: the Run Stamp omits persisted snapshot field(s) {missing}"
        )
    _require_value(presentation, "summary_metrics_source",
                   "sim_data.summary_statistics.metrics", pwhere)
    rows = _seq(presentation, "selected_rows", pwhere)
    actual = tuple(
        (_req_str(r, "key", pwhere), _req_str(r, "source", pwhere), bool(r.get("lookup_only")))
        for r in rows
    )
    expected = (
        ("selected_confidence_level", "inpSelectedConfidenceLevel", False),
        ("selected_quantile", "sim_data.summary_statistics", True),
        ("selected_contingency", "sim_data.contingency_ladder", True),
    )
    if actual != expected:
        raise SimContractError(
            f"{pwhere}: the selected reporting rows must be exactly {expected}, got {actual}. "
            "Selected Px and Contingency are LOOKUPS into the persisted ladders; computing "
            "either on the sheet would be a second engine."
        )
    _require_true(presentation, "blank_when_no_active_bank", pwhere)
    _require_true(presentation, "blank_when_selector_not_selectable", pwhere)


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
    if tuple(reserved) != LOCKED_RESERVED_ROWS:
        raise SimContractError(
            f"{where}: reserved_rows must be exactly the accepted tiling "
            f"{[list(r) for r in LOCKED_RESERVED_ROWS]}, got {[list(r) for r in reserved]}. "
            "The tiling IS the derivation of H, so its purposes are the audit trail for the "
            "technical ceiling, not labels."
        )

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
    actual_columns = tuple(
        (
            _req_str(c, "key", rwhere),
            _req_str(c, "column", rwhere),
            _req_str(c, "header", rwhere),
            _req_str(c, "value_type", rwhere),
        )
        for c in columns
    )
    if actual_columns != LOCKED_ITERATION_RECORD_COLUMNS:
        raise SimContractError(
            f"{rwhere}: the iteration columns must be exactly "
            f"{[list(c) for c in LOCKED_ITERATION_RECORD_COLUMNS]}, in order, got "
            f"{[list(c) for c in actual_columns]}. The contract already chose a deterministic "
            "machine layout; a column letter, header or type that can drift is one a "
            "materialiser and a reader can disagree about."
        )
    for _, column, _, _ in actual_columns:
        if not re.fullmatch(r"[A-Z]{1,3}", column):
            raise SimContractError(f"{rwhere}: column {column!r} is not a column letter")

    # THE SECOND BANK CONSUMES COLUMNS, NOT ROWS. Everything above - the tiling,
    # H, the header row and the first iteration row - has already been checked
    # against its accepted value, so a bank that tried to buy itself capacity by
    # moving the row axis has already failed.
    banks = _map(records, "banks", rwhere)
    bwhere = f"{rwhere}: banks"
    if tuple(banks) != LOCKED_BANK_LABELS:
        raise SimContractError(
            f"{bwhere}: the iteration banks must be exactly {list(LOCKED_BANK_LABELS)}, "
            f"in order, got {list(banks)}"
        )
    if {k: dict(v) for k, v in banks.items()} != LOCKED_ITERATION_BANKS:
        raise SimContractError(
            f"{bwhere}: the iteration bank columns must be exactly "
            f"{LOCKED_ITERATION_BANKS}, got {{k: dict(v) for k, v in banks.items()}}"
        )
    bank_a = dict(banks["A"])
    declared = {key: column for key, column, _h, _t in actual_columns}
    if bank_a != declared:
        raise SimContractError(
            f"{bwhere}: bank A must be the SAME layout the accepted `columns` block "
            f"declares, got {bank_a} against {declared}. Two spellings of one layout is "
            "two layouts."
        )
    if set(banks["A"].values()) & set(banks["B"].values()):
        raise SimContractError(
            f"{bwhere}: the two banks share a column; a candidate write would land in the "
            "published bank"
        )

    _validate_summary_statistics(block, where, header_row)
    _validate_contingency_ladder(block, where, header_row)

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
    # The COMPLETE record, not just key/row/group/type. Initials could be seeded
    # and enum owners swapped while every earlier check passed, which would have
    # let a materialiser write a partial successful snapshot into a workbook that
    # had never run.
    for key, expected in LOCKED_RUN_IDENTITY_COLUMNS.items():
        _require_value(identity, key, expected, iwhere)
    bank_columns = _map(identity, "bank_value_columns", iwhere)
    if dict(bank_columns) != LOCKED_RUN_IDENTITY_BANK_COLUMNS:
        raise SimContractError(
            f"{iwhere}: bank_value_columns must be exactly "
            f"{LOCKED_RUN_IDENTITY_BANK_COLUMNS}, got {dict(bank_columns)}"
        )
    if bank_columns["A"] != identity.get("value_column"):
        raise SimContractError(
            f"{iwhere}: bank A must reuse the shared value column, so a first-ever "
            "successful run lands exactly where the accepted single-bank layout put it"
        )
    if bank_columns["A"] == bank_columns["B"]:
        raise SimContractError(f"{iwhere}: the two banks share a value column")
    if identity.get("note_column") in set(bank_columns.values()):
        raise SimContractError(f"{iwhere}: the note column collides with a bank column")

    actual = tuple(
        (
            _req_str(f, "key", iwhere),
            _req_int(f, "row", iwhere),
            _req_str(f, "group", iwhere),
            _req_str(f, "label", iwhere),
            _req_str(f, "value_type", iwhere),
            f.get("enum"),
            f.get("initial"),
        )
        for f in fields
    )
    if actual != LOCKED_RUN_IDENTITY:
        differences = [
            f"{a[0]}: {a} != {e}"
            for a, e in zip(actual, LOCKED_RUN_IDENTITY)
            if a != e
        ]
        extra = [a[0] for a in actual if a[0] not in {r[0] for r in LOCKED_RUN_IDENTITY}]
        missing = [r[0] for r in LOCKED_RUN_IDENTITY if r[0] not in {a[0] for a in actual}]
        raise SimContractError(
            f"{iwhere}: the run-identity block must be exactly the accepted layout - key, row, "
            f"group, label, value type, enum owner and initial, in order. "
            f"unexpected={extra} missing={missing} differing={differences[:4]}. "
            "The persisted simulation identity is exact authority, not an extensible list, and "
            "every field that must be blank before its event must be explicitly blank."
        )

    # Cross-semantic agreement: the same fact must not be stated twice with two
    # different values. These live in different sections precisely because one is
    # the RULE and the other is the workbook cell that carries it.
    initials = {a[0]: a[6] for a in actual}
    if initials["next_auto_nonce"] != raw["seeding"]["nonce_lifecycle"]["initial"]:
        raise SimContractError(
            f"{iwhere}: next_auto_nonce.initial is {initials['next_auto_nonce']!r} but "
            f"seeding.nonce_lifecycle.initial is "
            f"{raw['seeding']['nonce_lifecycle']['initial']!r}"
        )
    if initials["last_run_id"] != raw["run_id"]["initial"]:
        raise SimContractError(
            f"{iwhere}: last_run_id.initial is {initials['last_run_id']!r} but run_id.initial "
            f"is {raw['run_id']['initial']!r}"
        )
    if initials["simulation_status"] is not None:
        raise SimContractError(
            f"{iwhere}: simulation_status.initial must be blank. A workbook that has never run "
            "must not present a derived status it never evaluated."
        )
    for key in ("run_id", "request_fingerprint", "result_digest", "effective_seed",
                "result_digest", "last_successful_stamp", "model_version"):
        if initials[key] is not None:
            raise SimContractError(
                f"{iwhere}: {key}.initial must be blank until a successful commit writes it; a "
                "seeded value would make a never-run workbook look like a partial success."
            )

    field_keys = [a[0] for a in actual]
    # `enum` is CONDITIONAL, and both directions are enforced: required when the
    # field is an enum, and refused when it is not.
    label_sets = raw.get("label_sets") or {}
    for entry in fields:
        key = entry.get("key")
        if entry.get("value_type") == "enum":
            name = _req_str(entry, "enum", iwhere)
            if name not in label_sets:
                raise SimContractError(
                    f"{iwhere}: field {key!r} names enum {name!r}, which is not declared in "
                    "label_sets"
                )
        elif "enum" in entry:
            raise SimContractError(
                f"{iwhere}: field {key!r} is not an enum but declares one. A label set that "
                "governs nothing is a semantic nobody reads."
            )

    excluded = tuple(block.get("excluded") or ())
    for required in LOCKED_SIM_DATA_EXCLUDED:
        if required not in excluded:
            raise SimContractError(
                f"{where}: excluded omits {required!r}. The accepted plan defers that "
                "retention; leaving it undeclared invites it back by accident."
            )

    _validate_phase7_records(block, where, header_row, first_row)

    return SimDataLayout(
        sheet=sheet,
        required_visibility=visibility,
        reserved_rows=tuple(reserved),
        header_row=header_row,
        first_iteration_row=first_row,
        footer_rows=0,
        reserved_row_count=h,
    )


LOCKED_SENSITIVITY_COLUMNS = (
    "driver_id", "driver_type", "driver_name", "rho", "abs_rho", "rank",
    "direction", "status",
)
"""The Sensitivity table, including the eighth field.

`status` is not decoration: it is where a zero-variance driver is reported as
"n/a - no variance" instead of being given a rho it does not have.
"""

# The columns the accepted iteration banks already own. A Phase-7 block that
# landed on any of them would overwrite published distribution data.
LOCKED_ITERATION_BANK_COLUMNS = frozenset({"B", "C", "D", "F", "G", "H"})


def _column_number(letter: str, where: str) -> int:
    text = str(letter).strip().upper()
    if not text or not text.isalpha():
        raise SimContractError(f"{where}: {letter!r} is not a column letter")
    value = 0
    for char in text:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def _phase7_block_shape(block: dict, name: str, where: str, header_row: int,
                        first_row: int) -> None:
    """Shared shape rules for both Phase-7 record blocks.

    THE POINT OF EVERY LINE HERE IS THAT PHASE 7 CONSUMES COLUMNS, NEVER ROWS.
    `reserved_rows` is the audit trail from which H, the first iteration row and
    the technical ceiling are derived, so a Phase-7 block that reserved a row
    would silently reduce the number of iterations the workbook can hold - and
    it would do so without touching the ceiling literal, which is exactly the
    kind of quiet capacity loss this contract is built to refuse.
    """
    bwhere = f"{where}: {name}"
    if _req_int(block, "header_row", bwhere) != header_row:
        raise SimContractError(
            f"{bwhere}: header_row must be the accepted iteration header row "
            f"{header_row}. A separate header row would be a reserved row, and "
            "reserved rows are what the technical ceiling is computed from."
        )
    if _req_int(block, "first_record_row", bwhere) != first_row:
        raise SimContractError(
            f"{bwhere}: first_record_row must be {first_row}, the accepted first "
            "iteration row, because this block shares that row axis."
        )
    if _req(block, "footer_rows", bwhere) != 0:
        raise SimContractError(f"{bwhere}: footer_rows must be 0")
    _require_true(block, "shares_row_axis_with_iteration_records", bwhere)
    _require_false(block, "consumes_reserved_rows", bwhere)


def _phase7_columns_clear_of_the_banks(letters, where: str) -> None:
    for letter in letters:
        if str(letter).strip().upper() in LOCKED_ITERATION_BANK_COLUMNS:
            raise SimContractError(
                f"{where}: column {letter!r} is owned by an iteration bank. A Phase-7 "
                "block written there would overwrite published distribution data."
            )


def _validate_phase7_records(block: dict, where: str, header_row: int,
                             first_row: int) -> None:
    swhere = f"{where}: sensitivity_records"
    sensitivity = _map(block, "sensitivity_records", swhere)
    _phase7_block_shape(sensitivity, "sensitivity_records", where, header_row, first_row)
    _require_value(sensitivity, "row_rule", LOCKED_SENSITIVITY_ROW_RULE, swhere)
    columns = _seq(sensitivity, "columns", swhere)
    actual = tuple(
        (_req_str(c, "key", swhere), _req_str(c, "column", swhere),
         _req_str(c, "header", swhere), _req_str(c, "value_type", swhere))
        for c in columns
    )
    _exact_sequence(actual, LOCKED_SENSITIVITY_COLUMN_LAYOUT, f"{swhere}: columns")
    _exact_sequence(tuple(row[0] for row in actual), LOCKED_SENSITIVITY_COLUMNS,
                    f"{swhere}: column keys")
    _phase7_columns_clear_of_the_banks([row[1] for row in actual], swhere)
    banks = _map(sensitivity, "banks", swhere)
    for bank in ("A", "B"):
        entry = _map(banks, bank, f"{swhere}: banks.{bank}")
        first = _req_str(entry, "first_column", f"{swhere}: banks.{bank}")
        last = _req_str(entry, "last_column", f"{swhere}: banks.{bank}")
        _phase7_columns_clear_of_the_banks((first, last), f"{swhere}: banks.{bank}")
        width = _column_number(last, swhere) - _column_number(first, swhere) + 1
        if width != len(LOCKED_SENSITIVITY_COLUMNS):
            raise SimContractError(
                f"{swhere}: banks.{bank} spans {width} columns but the table has "
                f"{len(LOCKED_SENSITIVITY_COLUMNS)} fields"
            )
    # THE STAMP shares the record columns and sits on rows the sheet already
    # reserves, above the first record row. It adds no row and moves no ceiling.
    stwhere = f"{swhere}: stamp"
    stamp = _map(sensitivity, "stamp", stwhere)
    _require_true(stamp, "published_written_last", stwhere)
    _require_true(stamp, "cleared_before_write", stwhere)
    _require_true(stamp, "surplus_rows_cleared", stwhere)
    stamp_columns = _map(stamp, "bank_value_columns", stwhere)
    for bank, expected in LOCKED_SENSITIVITY_STAMP_COLUMNS.items():
        _require_value(stamp_columns, bank, expected, stwhere)
    _phase7_columns_clear_of_the_banks(stamp_columns.values(), stwhere)
    stamp_fields = tuple(
        (_req_str(entry, "key", stwhere), _req_int(entry, "row", stwhere),
         _req_str(entry, "value_type", stwhere))
        for entry in _seq(stamp, "fields", stwhere))
    _exact_sequence(stamp_fields, LOCKED_SENSITIVITY_STAMP, f"{stwhere}: fields")
    assert stamp_fields[-1][0] == "published"
    for _key, row, _kind in stamp_fields:
        if row >= first_row:
            raise SimContractError(
                f"{stwhere}: the stamp field at row {row} is at or below the first "
                f"record row {first_row}; it would collide with a driver record")

    a_last = _column_number(_map(banks, "A", swhere)["last_column"], swhere)
    b_first = _column_number(_map(banks, "B", swhere)["first_column"], swhere)
    if b_first <= a_last:
        raise SimContractError(f"{swhere}: the sensitivity banks overlap")

    awhere = f"{where}: annual_records"
    annual = _map(block, "annual_records", awhere)
    _phase7_block_shape(annual, "annual_records", where, header_row, first_row)
    if _req_int(annual, "quantile_count", awhere) != LOCKED_ANNUAL_QUANTILE_COUNT:
        raise SimContractError(
            f"{awhere}: quantile_count must be {LOCKED_ANNUAL_QUANTILE_COUNT}, the "
            "accepted ladder width"
        )
    # THE LADDER, NEVER THE MATRIX.
    _require_false(annual, "iteration_level_annual_values_persisted", awhere)
    _require_value(annual, "row_rule", LOCKED_ANNUAL_ROW_RULE, awhere)
    _require_value(annual, "max_rows_owner",
                   "structure_contract.yaml: year_columns.max_generated_year_columns", awhere)
    _require_value(annual, "quantile_keys_owner",
                   "sim_data.summary_statistics.metrics", awhere)
    for group, locked in (("index_columns", LOCKED_ANNUAL_INDEX_COLUMNS),
                          ("quantile_first_column", LOCKED_ANNUAL_QUANTILE_FIRST_COLUMN),
                          ("selected_px_profile_columns", LOCKED_ANNUAL_PROFILE_COLUMNS)):
        mapping = _map(annual, group, f"{awhere}: {group}")
        for bank in ("A", "B"):
            gwhere = f"{awhere}: {group}.{bank}"
            entry = _map(mapping, bank, gwhere)
            for key, expected in locked[bank].items():
                _require_value(entry, key, expected, gwhere)
            _phase7_columns_clear_of_the_banks(entry.values(), gwhere)

    # THE LADDER MUST FIT BETWEEN ITS OWN START AND WHATEVER FOLLOWS IT. Eleven
    # quantiles occupy eleven columns; an allocation that overlaps the next
    # group would publish one number on top of another.
    for bank in ("A", "B"):
        starts = LOCKED_ANNUAL_QUANTILE_FIRST_COLUMN[bank]
        nominal = _column_number(starts["nominal"], awhere)
        pv = _column_number(starts["pv"], awhere)
        if pv - nominal < LOCKED_ANNUAL_QUANTILE_COUNT:
            raise SimContractError(
                f"{awhere}: bank {bank} allots {pv - nominal} columns to the nominal "
                f"ladder but it needs {LOCKED_ANNUAL_QUANTILE_COUNT}"
            )
        profile = _column_number(LOCKED_ANNUAL_PROFILE_COLUMNS[bank]["nominal"], awhere)
        if profile - pv < LOCKED_ANNUAL_QUANTILE_COUNT:
            raise SimContractError(
                f"{awhere}: bank {bank} allots {profile - pv} columns to the PV ladder "
                f"but it needs {LOCKED_ANNUAL_QUANTILE_COUNT}"
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


def _validate_publication_shell(sim: SimContract, spec: WorkbookSpec) -> None:
    """The workbook's Phase-6 shell must present exactly this contract's fields.

    `workbook.yaml` owns WHERE each label and formula sits; this contract owns
    WHICH fields exist. Two documents describing one layout drift the moment
    either moves, so the cross-check is the thing that makes the split safe.

    It also enforces the property the whole design rests on: every Results
    formula is a LOOKUP. A formula that averaged, deviated, interpolated a
    quantile or subtracted a contingency would be a second statistics engine on
    a worksheet, and no test of the VBA would ever see it.
    """
    shell = getattr(spec, "phase6_shell", None) or {}
    if not shell:
        return
    where = "workbook.yaml: phase6_shell"
    results = shell["results"]
    presentation = sim.raw["results_minimum"]["presentation"]

    declared = tuple(f["key"] for f in results["run_stamp"]["fields"])
    if declared != tuple(presentation["run_stamp_fields"]):
        raise SimContractError(
            f"{where}.results.run_stamp: presents {list(declared)}, but the contract's Run "
            f"Stamp is {list(presentation['run_stamp_fields'])}"
        )
    identity_rows = {f["key"]: f["row"] for f in sim.raw["sim_data"]["run_identity"]["fields"]}
    labels = {f["key"]: f["label"] for f in sim.raw["sim_data"]["run_identity"]["fields"]}
    for field_ in results["run_stamp"]["fields"]:
        if field_["key"] not in identity_rows:
            raise SimContractError(
                f"{where}.results.run_stamp: {field_['key']!r} is not a _SimData field"
            )
        if field_["label"] != labels[field_["key"]]:
            raise SimContractError(
                f"{where}.results.run_stamp: {field_['key']!r} is labelled "
                f"{field_['label']!r} but the contract calls it {labels[field_['key']]!r}"
            )
        if f"${identity_rows[field_['key']]}" not in field_["formula"]:
            raise SimContractError(
                f"{where}.results.run_stamp: the formula for {field_['key']!r} does not "
                f"read its _SimData row {identity_rows[field_['key']]}"
            )

    metrics = sim.raw["sim_data"]["summary_statistics"]["metrics"]
    shell_metrics = results["summary"]["metrics"]
    if tuple(m["key"] for m in shell_metrics) != tuple(m["key"] for m in metrics):
        raise SimContractError(
            f"{where}.results.summary: the presented metrics are not the persisted ones"
        )
    for shell_metric, metric in zip(shell_metrics, metrics):
        for measure in ("nominal", "pv"):
            if f"${metric['row']}" not in shell_metric[measure]:
                raise SimContractError(
                    f"{where}.results.summary: {metric['key']!r} does not read its "
                    f"persisted row {metric['row']}"
                )

    # EVERY formula, checked as one body of text.
    formulas = _shell_formulas(results)
    active = (f"{sim.raw['sim_data']['sheet']}!"
              f"${sim.raw['sim_data']['run_identity']['value_column']}"
              f"${identity_rows['active_bank']}")
    for formula in formulas:
        upper = formula.upper()
        for banned in LOCKED_RESULTS_FORBIDDEN_FUNCTIONS:
            if f"{banned}(" in upper:
                raise SimContractError(
                    f"{where}.results: a formula calls {banned}. Results presents numbers "
                    "the run persisted; it does not compute one."
                )
        if "-" in formula.split('"')[0] and "INDEX" not in formula:
            raise SimContractError(
                f"{where}.results: a formula subtracts. Contingency is persisted, not "
                "derived on the sheet."
            )
    # EVERY BANKED FORMULA READS THE SELECTOR. A formula pinned to one bank shows
    # the published distribution only half the time, and the half it gets wrong is
    # invisible until the second successful run.
    shared_keys = {key for key, group in
                   ((f["key"], f["group"]) for f in
                    sim.raw["sim_data"]["run_identity"]["fields"])
                   if group != "snapshot"}
    banked_formulas = [
        (f["key"], f["formula"]) for f in results["run_stamp"]["fields"]
        if f["key"] not in shared_keys
    ]
    banked_formulas.extend(
        (metric["key"], metric[measure])
        for metric in results["summary"]["metrics"] for measure in ("nominal", "pv")
    )
    for key, formula in banked_formulas:
        if active not in formula:
            raise SimContractError(
                f"{where}.results: the formula for {key!r} does not read the active-bank "
                "selector; a formula pinned to one bank shows the published distribution "
                "only half the time"
            )
    for key, formula in ((f["key"], f["formula"]) for f in results["run_stamp"]["fields"]
                         if f["key"] in shared_keys):
        if active in formula:
            raise SimContractError(
                f"{where}.results: {key!r} is SHARED state and must not be read through "
                "the bank selector"
            )
    selector = sim.raw["selected_confidence_level"]["source"]
    selected = results["selected"]
    for key in ("quantile_nominal", "quantile_pv", "contingency_nominal",
                "contingency_pv"):
        if selector not in selected[key] or "MATCH(" not in selected[key].upper():
            raise SimContractError(
                f"{where}.results.selected.{key}: must LOOK UP the selector in the "
                "persisted ladder"
            )


def _shell_formulas(results: dict[str, Any]) -> list[str]:
    out = [f["formula"] for f in results["run_stamp"]["fields"]]
    for metric in results["summary"]["metrics"]:
        out.extend([metric["nominal"], metric["pv"]])
    selected = results["selected"]
    out.extend([selected["confidence_level_formula"], selected["quantile_nominal"],
                selected["quantile_pv"], selected["contingency_nominal"],
                selected["contingency_pv"]])
    return out


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

    _validate_publication_shell(sim, spec)

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
