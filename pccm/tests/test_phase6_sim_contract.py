#!/usr/bin/env python3
"""PCCM Phase 6 Step-1 positive tests for the simulation contract.

`spec/sim_contract.yaml` is the sixth authority. It ENCODES what Phase-6 Step 0
settled; it does not choose it. These tests assert that what the contract says is
what the accepted Step-0 evidence says, value by value.

The evidence package is read here as a CONFORMANCE AUTHORITY, which is a test-time
relationship and only a test-time one: no production module imports
`evidence/`, and a separate test asserts that. The contract is the production
semantic authority after acceptance; the Step-0 package remains the record of WHY
those values were chosen.

NOTHING HERE SIMULATES. No RNG state advances, no uniform is generated, no jump
is executed, no variate is sampled, no iteration runs. Expected values come from
the retained vectors, never from an algorithm implemented to manufacture them.

Runs standalone or under pytest.
"""

from __future__ import annotations

import ast
import json
import sys
from pathlib import Path

import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    load_contract,
    load_driver_contract,
    load_sim_contract,
    load_spec,
    load_structure_contract,
    validate_sim_against,
)
from pccm_builder.sim_loader import (  # noqa: E402
    LOCKED_A1_P127,
    LOCKED_A2_P127,
    LOCKED_REQUEST_EFFECTIVE,
    LOCKED_REQUEST_FIELD_TYPES,
    LOCKED_REQUEST_GRAMMAR,
    LOCKED_REQUEST_RECORD_COUNT,
    LOCKED_RNG_CONSTANTS,
    LOCKED_SIM_STATES,
    MAX_EXCEL_ROWS,
)

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"
CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
DRIVERS_PATH = PCCM_ROOT / "spec" / "driver_contract.yaml"
STRUCTURE_PATH = PCCM_ROOT / "spec" / "structure_contract.yaml"
CALC_PATH = PCCM_ROOT / "spec" / "calc_contract.yaml"
SIM_PATH = PCCM_ROOT / "spec" / "sim_contract.yaml"
EVIDENCE = PCCM_ROOT / "evidence" / "phase6_step0"


def _sim():
    return load_sim_contract(SIM_PATH)


def _raw() -> dict:
    return yaml.safe_load(SIM_PATH.read_text(encoding="utf-8"))


def _evidence(relative: str) -> dict:
    return json.loads((EVIDENCE / relative).read_text(encoding="utf-8"))


# ===========================================================================
# A. the contract loads
# ===========================================================================
def test_01_sim_contract_loads() -> None:
    sim = _sim()
    assert sim.version == "1.0.0"
    assert sim.rng_version == 1
    assert sim.sim_method_version == 1


def test_02_cross_validates_against_the_other_five_authorities() -> None:
    validate_sim_against(
        _sim(),
        load_spec(SPEC_PATH),
        load_contract(CONTRACT_PATH),
        load_driver_contract(DRIVERS_PATH),
        load_structure_contract(STRUCTURE_PATH),
        yaml.safe_load(CALC_PATH.read_text(encoding="utf-8")),
    )


# ===========================================================================
# B. every authoritative constant is exact - checked against Step-0 evidence
# ===========================================================================
def test_03_rng_constants_match_the_retained_evidence() -> None:
    """The vectors file carries the constants that produced every retained uniform."""
    declared = _raw()["rng"]["constants"]
    retained = _evidence("vectors/rng_vectors.json")["constants"]
    for key in ("m1", "m2", "a12", "a13n", "a21", "a23n"):
        assert declared[key] == retained[key], key
    # `norm` is retained as a repr so the exact Double survives the JSON round trip.
    assert repr(declared["norm"]) == retained["norm"]
    assert declared == {**LOCKED_RNG_CONSTANTS}


def test_04_state_order_matches_the_retained_vectors() -> None:
    vectors = _evidence("vectors/jump_vectors.json")
    assert vectors["state_order_stored"].startswith("[s10,s11,s12,s20,s21,s22]")
    assert _raw()["rng"]["state"]["order"] == ["s10", "s11", "s12", "s20", "s21", "s22"]
    assert _raw()["rng"]["state"]["orientation"] == "oldest_first"
    assert _raw()["rng"]["state"]["matrix_operand_orientation"] == "newest_first"


def test_05_jump_matrices_match_the_retained_evidence_element_by_element() -> None:
    jump = _evidence("raw/jump.json")
    declared = _raw()["jump"]
    for name, retained in (
        ("a1_p127", jump["A1p127_derived_from_recurrence"]),
        ("a2_p127", jump["A2p127_derived_from_recurrence"]),
    ):
        assert declared[name] == retained, name
    assert declared["a1_p127"] == [list(r) for r in LOCKED_A1_P127]
    assert declared["a2_p127"] == [list(r) for r in LOCKED_A2_P127]


def test_06_jump_matrix_hashes_match_the_retained_evidence() -> None:
    jump = _evidence("raw/jump.json")
    declared = _raw()["jump"]
    assert declared["a1_p127_sha256"] == jump["A1p127_hash"]["sha256"]
    assert declared["a2_p127_sha256"] == jump["A2p127_hash"]["sha256"]


def test_07_jump_spacing_and_decomposition_are_the_settled_ones() -> None:
    declared = _raw()["jump"]
    assert declared["stream_spacing_exponent"] == 127
    assert declared["substream_spacing_exponent"] is None
    assert declared["substreams_used_in_phase_6"] is False
    assert declared["decomposition_h"] == 1 << 17
    assert declared["naive_floating_matrix_product_permitted"] is False


def test_08_auto_seed_mapping_matches_the_retained_seed_evidence() -> None:
    seed = _evidence("raw/seed_map.json")
    auto = _raw()["seeding"]["auto"]
    assert auto["modulus"] == seed["modulus"]
    assert auto["multiplier"] == seed["multiplier"]
    assert auto["period"] == seed["period"]
    assert auto["mapping_kind"] == "modular_exponentiation"
    assert auto["stepped_multiplication_is_the_authority"] is False


def test_09_nonce_lifecycle_matches_the_retained_evidence() -> None:
    life = _evidence("raw/seed_map.json")["nonce_lifecycle"]
    declared = _raw()["seeding"]["nonce_lifecycle"]
    assert declared["meaning"] == "next_nonce_to_allocate"
    assert life["auto_nonce_meaning"].startswith("THE NEXT NONCE")
    assert declared["initial"] == life["initial_persisted_value"]
    assert declared["failure_before_allocation_consumes_nonce"] is False
    assert declared["failure_after_allocation_consumes_nonce"] is True
    assert declared["order"][1] == "read_current_auto_nonce"
    assert declared["order"][3] == "persist_auto_nonce_plus_one"
    assert declared["order"].index("read_current_auto_nonce") < declared["order"].index(
        "persist_auto_nonce_plus_one"
    ) < declared["order"].index("begin_sampling")


def test_10_scalar_seed_expansion_matches_the_retained_seed_vectors() -> None:
    vectors = _evidence("vectors/seed_vectors.json")
    assert vectors["seed_to_state_rule"] == "state = [seed] * 6"
    scalar = _raw()["seeding"]["scalar_to_state"]
    assert scalar["rule"] == "repeated_scalar"
    assert scalar["mixer"] is None
    for example in vectors["examples"]:
        assert example["state"] == [example["seed"]] * 6


# ===========================================================================
# C. the exact Cheng formulation
# ===========================================================================
def test_11_cheng_orientations_match_the_retained_formulation() -> None:
    formulation = _evidence("raw/cheng_formulation.json")
    declared = _raw()["cheng"]
    assert formulation["BB"]["parameter_orientation"].startswith("a = min")
    assert formulation["BC"]["parameter_orientation"].startswith("a = max")
    assert declared["bb"]["orientation"]["a"] == "min(alpha0, beta0)"
    assert declared["bc"]["orientation"]["a"] == "max(alpha0, beta0)"


def test_12_cheng_literals_match_the_retained_formulation() -> None:
    formulation = _evidence("raw/cheng_formulation.json")
    declared = _raw()["cheng"]
    for branch, key in (("bb", "BB"), ("bc", "BC")):
        retained = [entry["value"] for entry in formulation[key]["literals"]]
        assert declared[branch]["literals"] == retained, branch


def test_13_the_squeeze_literals_are_literals_not_computed() -> None:
    """1.3862944 is those digits. It is NOT evaluated as log(4)."""
    formulation = _evidence("raw/cheng_formulation.json")
    entry = next(e for e in formulation["BB"]["literals"] if e["value"] == "1.3862944")
    assert "NOT computed as log(4)" in entry["kind"]
    assert _raw()["cheng"]["literals_are_literal"] is True
    assert _raw()["cheng"]["algebraic_simplification_permitted"] is False
    # SEMANTIC, not textual: `log(4)` may legitimately appear as documentation of
    # what the literal approximates. What must never happen is the literal being
    # written as a COMPUTATION in an expression the implementation follows.
    cheng = _raw()["cheng"]
    for branch in ("bb", "bc"):
        expressions = list(cheng[branch]["per_driver"]) + list(cheng[branch]["per_attempt"])
        assert any("1.3862944" in line for line in expressions), branch
        for line in expressions:
            assert "log(4)" not in line, f"{branch}: {line}"
            assert "log(5)" not in line, f"{branch}: {line}"
            assert "1/72" not in line and "7/9" not in line, f"{branch}: {line}"


def test_14_the_logit_form_is_the_locked_one() -> None:
    formulation = _evidence("raw/cheng_formulation.json")
    declared = _raw()["cheng"]
    assert declared["logit_form"] == formulation["logit_form"]["locked"]
    assert declared["logit_form_rejected_alternative"] == (
        formulation["logit_form"]["rejected_alternative"]
    )
    for branch in ("bb", "bc"):
        for line in declared[branch]["per_attempt"]:
            assert "log1p" not in line, f"{branch}: {line}"


def test_15_cheng_source_binding_matches_the_retained_evidence() -> None:
    formulation = _evidence("raw/cheng_formulation.json")
    binding = _raw()["cheng"]["source_binding"]
    assert binding["functions_sha256"] == formulation["source_binding"]["sha256"]


def test_16_cheng_vectors_are_a_conformance_authority_not_a_lookup_table() -> None:
    vectors = _evidence("vectors/cheng_vectors.json")
    declared = _raw()["cheng"]["conformance_vectors"]
    assert declared["role"] == "conformance_authority"
    assert declared["runtime_lookup_table"] is False
    assert vectors["both_dispatches_covered"] == {"BB": 3, "BC": 2}
    assert vectors["every_case_exercises_a_retry"] is True
    assert vectors["every_case_exercises_immediate_acceptance"] is True


def test_17_dispatch_boundary_belongs_to_bc_in_contract_and_evidence() -> None:
    formulation = _evidence("raw/cheng_formulation.json")
    assert formulation["dispatch"]["boundary"] == "equality belongs to BC"
    dispatch = _raw()["distributions"]["beta_pert"]["dispatch"]
    assert dispatch["equality_belongs_to"] == "BC"
    assert dispatch["comparison_operator"] == "strictly_greater_than"
    for case in _evidence("vectors/cheng_vectors.json")["cases"]:
        expected = "BB" if min(case["alpha"], case["beta"]) > 1.0 else "BC"
        assert case["dispatch"] == expected, case["label"]


def test_18_two_uniforms_per_non_degenerate_attempt() -> None:
    assert _raw()["cheng"]["uniforms_per_non_degenerate_proposal_attempt"] == 2
    for case in _evidence("vectors/cheng_vectors.json")["cases"]:
        for sample in case["samples"]:
            assert sample["uniforms_for_this_sample"] == (
                2 * sample["proposal_attempts_for_this_sample"]
            )


# ===========================================================================
# D. D6-18 and the degenerate rule
# ===========================================================================
def test_19_severity_invocation_is_unconditional() -> None:
    severity = _raw()["risk"]["severity"]
    assert severity["invocation_policy"] == "unconditional"
    assert severity["sampler_invoked_every_risk_iteration"] is True
    assert severity["value_used_only_when_occurred"] is True


def test_20_degenerate_consumes_nothing_in_contract_and_evidence() -> None:
    degenerate = _raw()["distributions"]["degenerate"]
    assert degenerate["uniforms_consumed"] == 0
    assert degenerate["sampler_entered"] is False
    assert degenerate["stream_state_changed"] is False
    assert degenerate["detected_before_parameterisation"] is True
    assert _raw()["risk"]["severity"]["degenerate_consumption"] == 0
    rows = _evidence("raw/degenerate_d6_18.json")["rows"]
    for row in (r for r in rows if r["degenerate"]):
        assert row["severity_uniforms_consumed"] == 0
        assert row["severity_stream_unchanged"] is True
        assert row["severity_sampler_invoked_every_iteration"] is True


def test_21_the_withdrawn_advancement_phrase_appears_in_no_contract_VALUE() -> None:
    """SEMANTIC, not textual.

    A comment quoting the withdrawn phrase in order to prohibit it is exactly
    what the file should contain. What must not exist is a contract VALUE saying
    it, so this walks the parsed document rather than the raw bytes - the same
    boundary the loader enforces.
    """
    values = []

    def walk(node):
        if isinstance(node, dict):
            for value in node.values():
                walk(value)
        elif isinstance(node, list):
            for value in node:
                walk(value)
        else:
            values.append(str(node).lower())

    walk(_raw())
    for value in values:
        assert "advances once per iteration" not in value, value


def test_22_occurrence_comparison_is_strictly_less_than() -> None:
    occurrence = _raw()["risk"]["occurrence"]
    assert occurrence["rule"] == "occurred = u_occurrence < probability"
    assert occurrence["comparison_operator"] == "strictly_less_than"
    assert occurrence["uniforms_per_risk_per_iteration"] == 1


# ===========================================================================
# E. identities and versions
# ===========================================================================
def test_23_request_fingerprint_preserves_the_phase5_prefix() -> None:
    block = _raw()["request_fingerprint"]
    assert block["section_order"] == ["HEADER", "COST", "RISK", "SIM"]
    assert block["section_order"][:3] == block["analytical_prefix"]
    assert block["existing_sections_modified"] is False


def test_24_sim_section_excludes_run_scoped_identities() -> None:
    sim = _raw()["request_fingerprint"]["sim_section"]
    for excluded in ("effective_seed", "auto_nonce", "run_id", "selected_confidence_level"):
        assert excluded in sim["excluded_fields"]
        assert excluded not in sim["fields"]
    assert sim["analytical_fingerprint_hashed_as_a_field"] is False


def test_25_result_digest_framing_matches_the_retained_digest_evidence() -> None:
    digest = _evidence("raw/digest.json")
    block = _raw()["result_digest"]
    assert block["stream_tag"] == digest["stream_tag"]
    assert block["section_name"] == digest["section_name"]
    assert block["record_fields"] == ["iteration_index", "total_nominal", "total_pv"]
    assert block["record_field_count"] == 3
    assert block["iteration_index_origin"] == 1
    assert block["samples_sorted_for_digest"] is False
    assert block["equality"] == "exact"
    assert "tolerance" not in block, "the tolerance key was removed entirely"


def test_26_digest_version_field_is_sim_method_version() -> None:
    assert _raw()["result_digest"]["version_field_source"] == "sim_method_version"
    assert _raw()["versions"]["result_digest_version_source"] == "sim_method_version"
    text = SIM_PATH.read_text(encoding="utf-8").lower()
    assert "result_digest_version:" not in text


def test_27_version_ownership_table_matches_the_retained_register() -> None:
    register = _evidence("raw/version_register.json")
    declared = _raw()["versions"]
    names = {v["name"] for v in register["versions"]}
    assert "RNG_VERSION" in names and "SIM_METHOD_VERSION" in names
    for version in register["versions"]:
        if version["name"] in ("RNG_VERSION", "SIM_METHOD_VERSION"):
            key = version["name"].lower()
            assert declared[key] == version["initial"], version["name"]
    # Every retained classification maps onto exactly one declared owner.
    owned = set(declared["bump_ownership"]["rng_version"]) | set(
        declared["bump_ownership"]["sim_method_version"]
    )
    assert len(owned) == len(declared["bump_ownership"]["rng_version"]) + len(
        declared["bump_ownership"]["sim_method_version"]
    ), "a change is owned by two versions"


# ===========================================================================
# F. D6-08 - derived, not declared
# ===========================================================================
def test_28_reserved_rows_tile_the_sheet_head_with_no_gap() -> None:
    reserved = _raw()["sim_data"]["reserved_rows"]
    expected = 1
    for entry in reserved:
        first, last = entry["rows"]
        assert first == expected, f"{entry['purpose']} starts at {first}, expected {expected}"
        expected = last + 1
    assert expected - 1 == _sim().reserved_rows_h


def test_29_d6_08_ceiling_is_derived_from_the_layout() -> None:
    sim = _sim()
    assert sim.reserved_rows_h == 33
    assert sim.layout.header_row == 33
    assert sim.layout.first_iteration_row == 34
    assert sim.layout.footer_rows == 0
    assert sim.max_iterations_representable == MAX_EXCEL_ROWS - 33 == 1048543
    ceiling = _raw()["iterations"]["technical_ceiling"]
    assert ceiling["reserved_rows_h"] == sim.reserved_rows_h
    assert ceiling["max_iterations_representable"] == sim.max_iterations_representable


def test_30_the_maximum_is_representable_and_one_more_is_not() -> None:
    """The boundary vectors D6-08 closes on.

    `max` occupies the last Excel row exactly; `max + 1` would need a row that
    does not exist, so it is a PRE-FLIGHT technical refusal - not a business
    validation, and not a COM error discovered mid-publish.
    """
    sim = _sim()
    first = sim.layout.first_iteration_row
    n_max = sim.max_iterations_representable
    assert first + n_max - 1 == MAX_EXCEL_ROWS
    assert first + (n_max + 1) - 1 == MAX_EXCEL_ROWS + 1 > MAX_EXCEL_ROWS


def test_31_the_technical_ceiling_is_not_a_business_rule() -> None:
    block = _raw()["iterations"]
    assert block["business_maximum"] is None
    assert "business_minimum" not in block
    assert block["business_minimum_owner"] == "input_contract.yaml"
    ceiling = block["technical_ceiling"]
    assert ceiling["refusal_kind"] == "technical"
    assert ceiling["presented_as_business_validation"] is False


def test_32_the_ceiling_refusal_precedes_auto_nonce_allocation() -> None:
    """A storage refusal must not consume an AUTO nonce."""
    ceiling = _raw()["iterations"]["technical_ceiling"]
    order = ceiling["refusal_precedes"]
    assert "auto_seed_allocation" in order
    assert "any_random_draw" in order
    assert ceiling["consumes_auto_nonce"] is False


def test_33_sim_data_stores_exactly_the_three_iteration_fields() -> None:
    records = _raw()["sim_data"]["iteration_records"]
    assert [c["key"] for c in records["columns"]] == [
        "iteration_index", "total_nominal", "total_pv",
    ]
    assert records["sorted"] is False
    assert records["order"] == "canonical_iteration_order"
    for excluded in ("per_driver_samples", "annual_stochastic_samples", "sensitivity_data"):
        assert excluded in _raw()["sim_data"]["excluded"]


def test_34_run_identity_carries_every_required_field() -> None:
    fields = {f["key"] for f in _raw()["sim_data"]["run_identity"]["fields"]}
    for required in (
        "run_id", "request_fingerprint", "result_digest", "seed_mode", "supplied_seed",
        "effective_seed", "consumed_auto_nonce", "iterations_run", "rng_version",
        "sim_method_version", "last_successful_stamp", "applied_timeline",
    ):
        assert required in fields, required


# ===========================================================================
# G. states, statistics, contingency, prerequisite
# ===========================================================================
def test_35_exactly_three_simulation_states() -> None:
    assert _raw()["sim_state"]["states"] == list(LOCKED_SIM_STATES)
    assert _raw()["label_sets"]["sim_state"] == list(LOCKED_SIM_STATES)
    assert _raw()["sim_state"]["no_success_valid_status"] is None


def test_36_a_failure_publishes_nothing_partial() -> None:
    failure = _raw()["sim_state"]["on_failure"]
    assert failure["prior_sim_data_preserved"] is True
    assert failure["prior_results_publication_preserved"] is True
    assert failure["partial_distribution_published"] is False


def test_37_phase5_prerequisite_is_current_and_never_recalculated() -> None:
    block = _raw()["prerequisite"]
    assert block["phase5_analytical_state_required"] == "CURRENT"
    assert block["silent_recalculation_permitted"] is False
    assert block["phase6_may_call_pccm_calculate"] is False


def test_38_percentile_is_type_7_with_convex_interpolation() -> None:
    pct = _raw()["statistics"]["percentile"]
    assert pct["method"] == "hyndman_fan_type_7"
    assert pct["formula"]["h"] == "(n - 1) * p"
    assert pct["formula"]["value"] == "(1 - f) * x[lo] + f * x[hi]"
    assert pct["interpolation"] == "convex"
    assert _raw()["statistics"]["standard_deviation"]["divisor"] == "n_minus_1"
    assert _raw()["statistics"]["sorting"] == "on_copies_only"


def test_39_contingency_baseline_is_the_deterministic_a() -> None:
    block = _raw()["contingency"]
    assert block["baseline"] == "deterministic_base_estimate_a"
    for forbidden in ("simulation_mean", "analytical_expected_total", "a_plus_emv"):
        assert forbidden in block["forbidden_baselines"]
    assert block["workbook_recommends_a_confidence_level"] is False


def test_40_run_id_is_a_success_counter() -> None:
    block = _raw()["run_id"]
    assert block["initial"] == 0
    assert block["first_successful_value"] == 1
    assert block["allocated_on"] == "successful_commit_only"
    assert block["failure_consumes"] is False
    assert block["maximum"] == 2147483647
    assert block["wrap_permitted"] is False


def test_41_results_scope_is_minimal_and_the_rest_is_deferred() -> None:
    block = _raw()["results_minimum"]
    assert block["sections"] == ["Run Stamp", "Summary Statistics"]
    for deferred in ("Annual Cash Flow", "Reconciliation presentation", "Dashboard",
                     "Charts", "Sensitivity"):
        assert deferred in block["deferred"]
    assert block["annual_simulated_samples_contracted"] is False


# ===========================================================================
# H. boundaries - what this contract must NOT own
# ===========================================================================
def test_42_the_seed_range_is_owned_by_the_input_contract() -> None:
    contract = load_contract(CONTRACT_PATH)
    validation = contract.inputs["random_seed"].validation
    assert validation["kind"] == "whole"
    assert validation["operator"] == "between"
    assert validation["formula1"] == "1"
    assert validation["formula2"] == "2147483646"
    assert contract.inputs["random_seed"].required is False
    assert contract.inputs["random_seed"].default is None


def test_43_the_sim_contract_carries_no_copy_of_the_seed_range() -> None:
    """`2147483646` may appear only as the AUTO PERIOD and the nonce exhaustion
    point - two facts about the nonce cycle, not about the input domain."""
    raw = _raw()
    allowed = {
        ("seeding", "auto", "period"),
        ("seeding", "nonce_lifecycle", "exhausted_value"),
    }
    found = set()

    def walk(node, prefix=()):
        if isinstance(node, dict):
            for key, value in node.items():
                walk(value, prefix + (key,))
        elif isinstance(node, list):
            for i, value in enumerate(node):
                walk(value, prefix + (i,))
        elif str(node) == "2147483646":
            found.add(prefix)

    walk(raw)
    assert found == allowed, f"unexpected seed-maximum occurrences: {sorted(found - allowed)}"


def test_44_the_monte_carlo_minimum_did_not_move() -> None:
    contract = load_contract(CONTRACT_PATH)
    validation = contract.inputs["monte_carlo_iterations"].validation
    assert validation["operator"] == "greaterThanOrEqual"
    assert validation["formula1"] == "1000"
    assert "1000" not in yaml.safe_dump(_raw()["iterations"])


def test_45_no_comparison_tolerance_lives_in_the_sim_contract() -> None:
    """ZERO tolerance semantics in the parsed contract - not even a null field.

    Comments may explain where the tolerance lives; the parsed document may not
    carry the semantic at all.
    """
    keys = []
    values = []

    def walk(node, prefix=()):
        if isinstance(node, dict):
            for key, value in node.items():
                keys.append(str(key).lower())
                walk(value, prefix + (key,))
        elif isinstance(node, list):
            for i, value in enumerate(node):
                walk(value, prefix + (i,))
        else:
            values.append(str(node).lower())

    walk(_raw())
    for key in keys:
        for token in ("tolerance", "ulp", "rel_tol", "abs_tol", "rtol", "atol", "epsilon"):
            assert token not in key, f"key {key!r} carries a tolerance semantic"
    for value in values:
        assert "tolerance" not in value, f"value {value!r} mentions a tolerance"


def test_46_every_authority_reference_resolves() -> None:
    sim = _sim()
    assert len(sim.authority_references) == 12
    owners = {r.owner for r in sim.authority_references}
    assert owners == {
        "input_contract.yaml", "driver_contract.yaml", "structure_contract.yaml",
        "calc_contract.yaml", "workbook.yaml",
    }


# ===========================================================================
# I. scope discipline - Step 1 implements no algorithm
# ===========================================================================
SIM_LOADER = PCCM_ROOT / "builder" / "pccm_builder" / "sim_loader.py"


def test_47_the_loader_implements_no_rng_or_sampler() -> None:
    """Step 1 is the contract step. No production function may advance a state,
    generate a uniform, execute a jump, sample a variate or run an iteration."""
    source = SIM_LOADER.read_text(encoding="utf-8")
    banned = (
        "def next_u", "def sample", "def advance", "def jump_state", "def cheng",
        "math.log", "math.exp", "math.sqrt", "random.", "import random",
    )
    for token in banned:
        assert token not in source, f"sim_loader.py contains {token!r}"


def test_48_no_production_module_reads_the_evidence_package() -> None:
    """Tests may read retained evidence for conformance. Production must not.

    SEMANTIC, not a substring scan. The loader legitimately holds the DECLARED
    evidence paths as locked literals so a contract that claims to bind evidence
    cannot point at "banana" - that is a comparison target, not a dependency.
    What must not exist is production code that OPENS anything under `evidence/`,
    or imports from it.
    """
    builder = PCCM_ROOT / "builder"
    readers = ("open", "read_text", "read_bytes", "load", "loads", "glob", "rglob")
    offenders = []
    for path in sorted(builder.rglob("*.py")):
        text = path.read_text(encoding="utf-8")
        tree = ast.parse(text)
        for node in ast.walk(tree):
            if isinstance(node, (ast.Import, ast.ImportFrom)):
                names = [a.name for a in getattr(node, "names", [])]
                if getattr(node, "module", None):
                    names.append(node.module)
                for name in names:
                    if "evidence" in (name or "") or "phase6_step0" in (name or ""):
                        offenders.append(f"{path.name}: imports {name}")
            if isinstance(node, ast.Call):
                func = node.func
                name = getattr(func, "attr", None) or getattr(func, "id", None)
                if name in readers:
                    literals = " ".join(
                        a.value for a in ast.walk(node)
                        if isinstance(a, ast.Constant) and isinstance(a.value, str)
                    )
                    if "evidence" in literals or "phase6_step0" in literals:
                        offenders.append(f"{path.name}: {name}() over an evidence path")
    assert not offenders, f"production reads the evidence package: {offenders}"


def test_48b_the_declared_evidence_paths_resolve_to_the_retained_artefacts() -> None:
    """The contract's OWN declared paths are followed and checked.

    A contract that claims to bind evidence but can point at "banana" is not
    bound. This resolves what the contract declares - not what the test assumes -
    and verifies the artefacts and the hash behind them.
    """
    cheng = _raw()["cheng"]
    for declared in (
        cheng["source_binding"]["evidence_file"],
        cheng["conformance_vectors"]["evidence_file"],
    ):
        assert declared.startswith("evidence/phase6_step0/"), declared
        resolved = PCCM_ROOT / declared
        assert resolved.is_file(), f"{declared} does not resolve to a retained artefact"

    formulation = json.loads(
        (PCCM_ROOT / cheng["source_binding"]["evidence_file"]).read_text(encoding="utf-8")
    )
    assert cheng["source_binding"]["functions_sha256"] == (
        formulation["source_binding"]["sha256"]
    )
    vectors = json.loads(
        (PCCM_ROOT / cheng["conformance_vectors"]["evidence_file"]).read_text(encoding="utf-8")
    )
    assert vectors["both_dispatches_covered"] == {"BB": 3, "BC": 2}

    jump = _evidence("raw/jump.json")
    declared_jump = _raw()["jump"]
    assert declared_jump["a1_p127_sha256"] == jump["A1p127_hash"]["sha256"]
    assert declared_jump["a2_p127_sha256"] == jump["A2p127_hash"]["sha256"]
    # And the hashes actually hash the declared matrices.
    import hashlib

    for name, key in (("a1_p127", "a1_p127_sha256"), ("a2_p127", "a2_p127_sha256")):
        text = ";".join(",".join(str(v) for v in row) for row in declared_jump[name])
        assert hashlib.sha256(text.encode("ascii")).hexdigest() == declared_jump[key], name


def test_49_the_loader_parses_but_never_evaluates_contract_expressions() -> None:
    """The contract's formulas are TEXT. Nothing eval()s or compiles them."""
    source = SIM_LOADER.read_text(encoding="utf-8")
    tree = ast.parse(source)
    called = {
        node.func.id
        for node in ast.walk(tree)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
    }
    for banned in ("eval", "exec", "compile"):
        assert banned not in called, f"sim_loader.py calls {banned}()"


def test_50_only_the_authorised_phase6_vba_exists() -> None:
    """Step 1 authorised no Phase-6 VBA. Steps 6 and 7 authorised one module each.

    The rule has not loosened: a Phase-6 module still may not appear without a
    step that authorises it, and the remaining four are still absent.
    `modSimContract.bas` is still not in `src/vba` - it is GENERATED, and a
    hand-written copy would be a second definition of every literal it projects.
    """
    src = PCCM_ROOT / "src" / "vba"
    names = {p.name for p in src.glob("*.bas")} | {p.name for p in src.glob("*.cls")}
    for authorised in ("modSimRng.bas", "modSimSample.bas", "modSimEngine.bas",
                       "modSimStats.bas"):
        assert authorised in names, authorised
    for banned in ("modSimFingerprint.bas", "modSimReport.bas", "modSimContract.bas"):
        assert banned not in names, f"{banned} exists; no step authorises it there"
    assert "modSimRng.bas" in names
    for path in sorted(src.glob("*.bas")):
        if path.stem == "modSimRng":
            continue
        text = path.read_text(encoding="utf-8", errors="replace")
        assert "MRG32k3a" not in text, f"{path.name} references MRG32k3a"


# ===========================================================================
# J. the corrected state authority - full truth table
# ===========================================================================
def test_51_the_state_derivation_is_total_and_mutually_exclusive() -> None:
    """Every combination lands in exactly one outcome, and none lands nowhere.

    Revision 6's predicates had a hole - a corrected-then-restored request was
    none of the three - and an overlap. The corrected rules are ordered, so
    exclusivity is structural and totality is checked here by exhaustion.
    """
    from pccm_builder.sim_loader import derive_sim_status

    outcomes = {}
    for prereq in (True, False):
        for snapshot in (True, False):
            for matches in (True, False):
                outcomes[(prereq, snapshot, matches)] = derive_sim_status(
                    prereq, snapshot, matches
                )
    assert len(outcomes) == 8
    assert set(outcomes.values()) == {"INVALID", "CURRENT", "STALE", None}
    for key, value in outcomes.items():
        assert value in ("INVALID", "CURRENT", "STALE", None), key


def test_52_the_required_state_cases_resolve_as_the_correction_specifies() -> None:
    """A .. F from the Step-1 review, one assertion each."""
    from pccm_builder.sim_loader import derive_sim_status

    # A. success A -> REFUSED invalid edit -> restore A. Attempt is REFUSED and
    #    must not matter: the request matches, so the result is CURRENT.
    assert derive_sim_status(True, True, True) == "CURRENT"
    # B. success A -> valid changed request B.
    assert derive_sim_status(True, True, False) == "STALE"
    # C. success A -> FAILED on B, rolled back. Viewing B is STALE; restored A
    #    is CURRENT. The FAILED attempt changes neither.
    assert derive_sim_status(True, True, False) == "STALE"
    assert derive_sim_status(True, True, True) == "CURRENT"
    # D. current prerequisites invalid, whatever the history.
    for snapshot in (True, False):
        for matches in (True, False):
            assert derive_sim_status(False, snapshot, matches) == "INVALID"
    # E. no successful simulation, current request valid -> BLANK.
    assert derive_sim_status(True, False, True) is None
    assert derive_sim_status(True, False, False) is None
    # F. no successful simulation, current request invalid -> INVALID.
    assert derive_sim_status(False, False, True) == "INVALID"


def test_53_attempt_history_cannot_change_the_derived_status() -> None:
    """The derivation takes no attempt argument at all - it CANNOT read it.

    That is stronger than asserting the outcome is unchanged for each attempt
    label: the parameter does not exist, so no future edit can quietly add a
    branch on it without changing the signature this test pins.
    """
    import inspect

    from pccm_builder.sim_loader import derive_sim_status

    parameters = list(inspect.signature(derive_sim_status).parameters)
    assert parameters == [
        "prerequisites_resolve",
        "successful_snapshot_exists",
        "request_fingerprint_matches",
    ]
    for token in ("attempt", "refused", "failed"):
        assert token not in inspect.getsource(derive_sim_status).lower(), token

    contract = _raw()["sim_state"]
    assert contract["attempt_result_participates_in_derivation"] is False
    assert contract["attempt_axis_is_orthogonal"] is True
    for rule in contract["derivation"]["rules"]:
        for token in ("attempt", "refused", "failed"):
            assert token not in rule["condition"].lower(), rule


def test_54_the_contract_derivation_matches_the_function() -> None:
    """The YAML rules and the callable must not drift apart."""
    from pccm_builder.sim_loader import derive_sim_status

    rules = _raw()["sim_state"]["derivation"]["rules"]
    assert [r["order"] for r in rules] == [1, 2, 3, 4]
    assert _raw()["sim_state"]["derivation"]["ordered"] is True
    by_status = {r["order"]: r["status"] for r in rules}
    assert by_status[1] == "INVALID" and derive_sim_status(False, True, True) == "INVALID"
    assert by_status[2] is None and derive_sim_status(True, False, True) is None
    assert by_status[3] == "CURRENT" and derive_sim_status(True, True, True) == "CURRENT"
    assert by_status[4] == "STALE" and derive_sim_status(True, True, False) == "STALE"


# ===========================================================================
# K. the contribution contract
# ===========================================================================
def test_55_cost_line_samples_unit_cost_with_quantity_outside() -> None:
    cost = _raw()["contribution"]["cost_line"]
    assert cost["sampled_quantity"] == "unit_cost"
    assert cost["total_cost_uncertainty_sampled"] is False
    assert cost["quantity_inside_distribution"] is False
    assert cost["quantity_is_deterministic"] is True
    assert cost["quantity_applications"] == 1
    assert cost["probability_applies"] is False
    assert cost["nominal"] == "unit_cost * Quantity * Knom"
    assert cost["pv"] == "unit_cost * Quantity * Kpv"


def test_56_risk_contributes_severity_without_quantity() -> None:
    risk = _raw()["contribution"]["risk"]
    assert risk["quantity_applies"] is False
    assert risk["probability_folded_into_k_factors"] is False
    assert risk["occurrence_and_severity_share_a_stream"] is False
    assert risk["nominal_when_occurred"] == "severity * Knom"
    assert risk["pv_when_occurred"] == "severity * Kpv"
    assert risk["nominal_when_not_occurred"] == 0
    assert risk["pv_when_not_occurred"] == 0
    assert risk["occurred"] == "occurrence_uniform < Probability"


def test_57_pv_is_an_independent_contribution_not_a_discounted_nominal() -> None:
    assert _raw()["contribution"]["pv_derived_from_nominal"] is False
    assert _raw()["contribution"]["iteration_total"]["measures_independent"] is True
    assert _raw()["contribution"]["iteration_total"]["order_source"] == "accumulation"


# ===========================================================================
# L. kernel, numerical domain, dependence, publication
# ===========================================================================
def test_58_the_hot_loop_touches_no_worksheet_or_com_object() -> None:
    kernel = _raw()["kernel"]
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
        assert kernel[flag] is False, flag
    assert kernel["inputs_resolved_once_before_simulation"] is True
    for resolved in ("knom_per_driver", "kpv_per_driver", "quantities", "probabilities"):
        assert resolved in kernel["resolved_before_loop"], resolved


def test_59_the_phase5_numerical_domain_is_inherited_unnarrowed() -> None:
    domain = _raw()["numerical_domain"]
    assert domain["negative_values_legal"] is True
    assert domain["supports_crossing_zero_legal"] is True
    assert domain["positivity_rule"] is None
    assert domain["magnitude_restriction"] is None
    assert domain["narrower_than_phase5"] is False
    assert domain["representable_result_refused_for_naive_intermediate_overflow"] is False
    assert domain["silent_non_finite_result_permitted"] is False
    assert domain["disciplines"]["accumulation"] == "accepted_safe_signed_sum"
    assert domain["disciplines"]["percentile_interpolation"] == "convex"


def test_60_drivers_are_sampled_independently() -> None:
    dep = _raw()["dependence"]
    assert dep["inter_driver_dependence"] == "independent"
    assert dep["correlation_matrix_supported"] is False
    assert dep["copula_supported"] is False
    assert dep["shared_or_hidden_dependence_permitted"] is False


def test_61_publication_is_commit_last_and_results_never_recomputes() -> None:
    pub = _raw()["publication"]
    assert pub["persisted_source_of_truth"] == "_SimData"
    assert pub["results_derives_from"] == "_SimData"
    assert pub["results_recomputes_monte_carlo"] is False
    assert pub["commit_last"] is True
    assert pub["partial_new_distribution_published_on_refusal_or_failure"] is False
    assert pub["prior_successful_publication_survives"] is True


def test_62_phase6_adds_no_user_surface_and_no_cancellation() -> None:
    surface = _raw()["command_surface"]
    assert surface["automation_endpoint"] == "PCCM_RunSimulation"
    for flag in ("user_facing_run_button_in_phase_6", "msgbox_introduced_by_phase_6",
                 "userform_introduced_by_phase_6", "ribbon_introduced_by_phase_6",
                 "read_accessor_names_settled"):
        assert surface[flag] is False, flag
    assert _raw()["interruption"]["user_cancellation_supported_in_phase_6"] is False


# ===========================================================================
# M. the ladder, the run stamp, and content-bound authority
# ===========================================================================
def test_63_every_selectable_percentile_is_retained_by_reference() -> None:
    """Resolved from the OWNER. No ladder values are copied into sim_contract."""
    from pccm_builder.sim_loader import retained_percentiles

    document = yaml.safe_load(CONTRACT_PATH.read_text(encoding="utf-8"))
    ladder = next(
        t for t in document["config_tables"] if t["key"] == "confidence_levels"
    )["values"]
    retained = retained_percentiles(_sim(), document)

    assert retained[0] == "P10", "P10 is the fixed headline"
    for value in ladder:
        assert value in retained, f"{value} is selectable but would not be stored"
    assert len(retained) == len(ladder) + 1 == 11
    assert set(retained) == {"P10"} | set(ladder)

    # The values must NOT appear in the simulation contract itself.
    text = SIM_PATH.read_text(encoding="utf-8")
    for value in ladder:
        if value in ("P50", "P70", "P90"):
            continue  # legitimately named as headline percentiles
        assert value not in text, f"{value} was copied into sim_contract"


def test_64_model_version_is_persisted_in_the_run_stamp() -> None:
    fields = {f["key"]: f for f in _raw()["sim_data"]["run_identity"]["fields"]}
    assert "model_version" in fields, "the Run Stamp requires a model version"
    assert fields["model_version"]["group"] == "snapshot", (
        "the model version at the time of the successful run is snapshot data, "
        "not a live lookup when Results is displayed"
    )
    assert fields["model_version"]["value_type"] == "text"
    document = yaml.safe_load(SPEC_PATH.read_text(encoding="utf-8"))
    assert str(document["model"]["model_version"]).strip()
    concepts = {r.concept for r in _sim().authority_references}
    assert "model version" in concepts


def test_65_the_authority_bindings_check_content_not_only_resolution() -> None:
    """Both bindings that were false before: visibility and family membership."""
    document = yaml.safe_load(SPEC_PATH.read_text(encoding="utf-8"))
    sheet = next(s for s in document["sheets"] if s["name"] == "_SimData")
    assert sheet["visibility"] == _raw()["sim_data"]["required_visibility"] == "veryHidden"

    results = next(s for s in document["sheets"] if s["name"] == "Results")
    titles = {b.get("title") for b in results["blocks"]}
    for section in _raw()["results_minimum"]["sections"]:
        assert section in titles, section

    inputs = yaml.safe_load(CONTRACT_PATH.read_text(encoding="utf-8"))
    master = next(t for t in inputs["config_tables"] if t["key"] == "distributions")
    assert set(master["values"]) == set(_raw()["distributions"]["families"])


def test_66_the_run_identity_layout_is_exact() -> None:
    from pccm_builder.sim_loader import LOCKED_RUN_IDENTITY

    identity = _raw()["sim_data"]["run_identity"]
    fields = identity["fields"]
    actual = tuple(
        (f["key"], f["row"], f["group"], f["label"], f["value_type"],
         f.get("enum"), f.get("initial"))
        for f in fields
    )
    assert actual == LOCKED_RUN_IDENTITY
    assert len(actual) == 22
    assert [f["row"] for f in fields] == list(range(8, 30))
    assert (identity["label_column"], identity["value_column"], identity["note_column"]) == (
        "B", "D", "F",
    )

    # Cross-semantic initials and enum ownership.
    by_key = {f["key"]: f for f in fields}
    assert by_key["next_auto_nonce"]["initial"] == 0 == (
        _raw()["seeding"]["nonce_lifecycle"]["initial"]
    )
    assert by_key["last_run_id"]["initial"] == 0 == _raw()["run_id"]["initial"]
    assert by_key["last_attempt_result"]["initial"] == "NONE"
    assert by_key["simulation_status"]["initial"] is None
    assert by_key["seed_mode"]["enum"] == "seed_mode"
    assert by_key["last_attempt_result"]["enum"] == "attempt_result"
    assert by_key["last_attempt_seed_mode"]["enum"] == "seed_mode"
    assert by_key["simulation_status"]["enum"] == "sim_state"
    for field in fields:
        assert ("enum" in field) == (field["value_type"] == "enum"), field["key"]


def test_66b_the_iteration_columns_are_exact() -> None:
    from pccm_builder.sim_loader import LOCKED_ITERATION_RECORD_COLUMNS

    columns = _raw()["sim_data"]["iteration_records"]["columns"]
    actual = tuple(
        (c["key"], c["column"], c["header"], c["value_type"]) for c in columns
    )
    assert actual == LOCKED_ITERATION_RECORD_COLUMNS
    assert actual == (
        ("iteration_index", "B", "Iteration", "integer"),
        ("total_nominal", "C", "Total Nominal", "double"),
        ("total_pv", "D", "Total PV", "double"),
    )


def test_67_d6_11_activated_exactly_once_and_only_with_its_owner() -> None:
    """Step 1 wrote the precondition; Step 6 is the commit that satisfied it.

    The precondition is unchanged and still recorded. What changed is that its
    owner now exists, so exactly one grant has been made - to the module that
    owns the construct, and to nothing else.
    """
    structure = load_structure_contract(STRUCTURE_PATH)
    scoped = [r for r in structure.forbidden_construct_rules if r.is_scoped]
    assert [(r.construct, tuple(r.allowed_in)) for r in scoped] == [
        ("MRG32k3a", ("modSimRng",))
    ], scoped
    declared = {m.name for m in structure.vba_modules}
    for rule in scoped:
        for owner in rule.allowed_in:
            assert owner in declared, f"{owner} is not a declared module"
    assert (PCCM_ROOT / "src" / "vba" / "modSimRng.bas").is_file(), (
        "the grant landed without its owner"
    )
    record = (PCCM_ROOT / "docs" / "phase6_step1.md").read_text(encoding="utf-8")
    assert "activation precondition" in record.lower()
    assert "forbidden_in" in record


# ===========================================================================
# N. Uniform degeneracy - the four required conformance cases
# ===========================================================================
def _degenerate_by_contract(family: str, a, m, b) -> bool:
    """Evaluate the CONTRACT's declared condition for a family.

    The predicate is read from the contract rather than restated here, so this
    tests what the document says and not what the test author remembered.
    """
    condition = _raw()["distributions"]["degenerate"]["conditions"][family]
    if condition == "a == b":
        return a == b
    if condition == "a == m == b":
        return a == m == b
    raise AssertionError(f"unrecognised degeneracy condition {condition!r}")


def test_68_uniform_degeneracy_is_family_specific() -> None:
    conditions = _raw()["distributions"]["degenerate"]["conditions"]
    assert conditions["uniform"] == "a == b"
    assert conditions["triangular"] == "a == m == b"
    assert conditions["beta_pert"] == "a == m == b"
    assert _raw()["distributions"]["degenerate"]["most_likely_read_by_uniform_degeneracy"] is False
    assert _raw()["distributions"]["uniform"]["most_likely_used"] is False
    assert _raw()["distributions"]["uniform"]["most_likely_affects_degeneracy"] is False
    assert _raw()["distributions"]["uniform"]["most_likely_affects_uniform_consumption"] is False


def test_69_the_four_required_uniform_cases() -> None:
    """A .. D from the Step-1 hardening review.

    Accepted Phase-5 D1 ignores Uniform's Most Likely numerically and excludes it
    from the calculation fingerprint. It must therefore not be able to change
    dispatch, RNG consumption, stream state or request identity.
    """
    degenerate = _raw()["distributions"]["degenerate"]
    per_sample = _raw()["distributions"]["uniform"]["uniforms_per_non_degenerate_sample"]

    # A. Min = Max, Most Likely BLANK -> degenerate, zero consumption.
    assert _degenerate_by_contract("uniform", 100.0, None, 100.0)
    # B. Min = Max, Most Likely POPULATED and unrelated -> still degenerate.
    assert _degenerate_by_contract("uniform", 100.0, 42.0, 100.0)
    # Under the withdrawn common predicate, case B was NOT degenerate: it would
    # have entered the sampler and consumed a uniform.
    assert not (100.0 == 42.0 == 100.0)

    for case in (None, 42.0, 100.0, -7.5):
        assert _degenerate_by_contract("uniform", 100.0, case, 100.0), case
        assert degenerate["uniforms_consumed"] == 0
        assert degenerate["sampler_entered"] is False
        assert degenerate["stream_state_changed"] is False

    # C. Two different ignored Most Likely values, same Min/Max: identical
    #    semantics, so identical consumption.
    assert _degenerate_by_contract("uniform", 5.0, 1.0, 5.0) == (
        _degenerate_by_contract("uniform", 5.0, 999.0, 5.0)
    )

    # D. Non-degenerate Uniform a < b -> exactly one uniform.
    assert not _degenerate_by_contract("uniform", 0.0, 50.0, 100.0)
    assert not _degenerate_by_contract("uniform", 0.0, None, 100.0)
    assert per_sample == 1

    # Triangular and Beta-PERT keep the three-way condition, and under the
    # accepted ordering a <= m <= b it is equivalent for legal input.
    for family in ("triangular", "beta_pert"):
        assert _degenerate_by_contract(family, 7.0, 7.0, 7.0)
        assert not _degenerate_by_contract(family, 0.0, 50.0, 100.0)


# ===========================================================================
# Step-10A - the request-fingerprint grammar closure
#
# Step 0 locked the SIM extension's SEMANTIC fields and their order and stopped
# there. Several byte-distinct streams satisfied that: F_I or F_N for iterations,
# one record or five, an AUTO seed omitted or blank or zero, versions as integers
# or as text. `result_digest` had token-level authority from the start; these
# tests are the request fingerprint being brought to the same standard.
# ===========================================================================
def _request_section() -> dict:
    return _raw()["request_fingerprint"]["sim_section"]


def test_70_the_sim_extension_is_exactly_one_record() -> None:
    section = _request_section()
    assert section["name"] == "SIM"
    assert section["record_count"] == LOCKED_REQUEST_RECORD_COUNT == 1
    # Five one-field records would carry the same semantics as different bytes.
    assert section["record_count"] != len(section["fields"])


def test_71_every_field_has_exactly_one_canonical_encoder() -> None:
    types = _request_section()["field_types"]
    assert types == LOCKED_REQUEST_FIELD_TYPES
    assert types == {
        "iterations": "F_I",
        "seed_mode": "F_S",
        "supplied_seed": "F_I",
        "rng_version": "F_I",
        "sim_method_version": "F_I",
    }
    # No integer identity is a Double. A count, a seed and a version are
    # structural facts; F_N would let a version of 1 collide with a Double of 1.
    assert "F_N" not in set(types.values())
    assert set(types) == set(_request_section()["fields"])


def test_72_the_field_names_are_semantic_position_and_are_never_encoded() -> None:
    section = _request_section()
    assert section["encoded_field_names"] is False
    for production in section["grammar"].values():
        for name in section["fields"]:
            assert f'"{name}"' not in production, name


def test_73_auto_and_fixed_have_different_record_shapes() -> None:
    effective = _request_section()["effective_records"]
    assert list(effective) == ["AUTO", "FIXED"]
    assert effective["AUTO"]["field_count"] == 4
    assert effective["FIXED"]["field_count"] == 5
    assert effective["AUTO"]["fields"] == [
        "iterations", "seed_mode", "rng_version", "sim_method_version"]
    assert effective["FIXED"]["fields"] == [
        "iterations", "seed_mode", "supplied_seed", "rng_version", "sim_method_version"]
    for shape in effective.values():
        assert shape["field_count"] == len(shape["fields"])
    assert effective == {
        mode: {"field_count": len(fields), "fields": list(fields)}
        for mode, fields in LOCKED_REQUEST_EFFECTIVE.items()
    }


def test_74_an_auto_supplied_seed_is_absent_and_not_a_sentinel() -> None:
    section = _request_section()
    assert "supplied_seed" not in section["effective_records"]["AUTO"]["fields"]
    assert section["auto_supplied_seed_representation"] == "absent"
    assert section["supplied_seed_present_only_when"] == "FIXED"
    # Not zero, not blank, not null, not the previous effective seed.
    assert "supplied_seed" not in _request_section()["grammar"]["auto_record"]
    assert "I1:0" not in _request_section()["grammar"]["auto_record"]
    assert _raw()["request_fingerprint"]["auto_blank_seed_remains_recomputable"] is True


def test_75_the_grammar_is_locked_token_by_token() -> None:
    grammar = _request_section()["grammar"]
    assert grammar == LOCKED_REQUEST_GRAMMAR
    assert grammar["section"] == 'F_S("SIM") F_I(1) sim_record'
    assert grammar["auto_record"] == (
        'F_I(4) F_I(iterations) F_S("AUTO") F_I(rng_version) F_I(sim_method_version)')
    assert grammar["fixed_record"] == (
        'F_I(5) F_I(iterations) F_S("FIXED") F_I(supplied_seed) F_I(rng_version) '
        'F_I(sim_method_version)')
    # The same standard the result digest has always carried.
    assert set(_raw()["result_digest"]["grammar"]) == {"stream", "section", "record"}


def test_76_the_extension_carries_no_stream_tag_and_no_stream_version() -> None:
    section = _request_section()
    assert section["stream_tag_repeated_in_extension"] is False
    assert section["stream_version_repeated_in_extension"] is False
    assert section["stream_tag_owner"] == "calc_contract.yaml"
    for production in section["grammar"].values():
        for banned in ("PCCM-FP", "FP_VERSION", "SIM_FP_VERSION", "REQUEST_FP_VERSION"):
            assert banned not in production, banned
    # No invented stream version exists anywhere in the LOADED contract, at any
    # depth - keys or values. The document is allowed to say in prose that it
    # refuses to carry one; a YAML comment is not a key.
    def walk(node):
        if isinstance(node, dict):
            for key, value in node.items():
                yield str(key)
                yield from walk(value)
        elif isinstance(node, list):
            for item in node:
                yield from walk(item)
        elif isinstance(node, str):
            yield node

    tokens = list(walk(_raw()))
    for invented in ("SIM_FP_VERSION", "REQUEST_FP_VERSION", "sim_fp_version",
                     "request_fp_version"):
        assert not any(invented in token for token in tokens), invented
    assert not any("PCCM-FP" in token for token in walk(_request_section()))
    # rng_version and sim_method_version are FIELDS INSIDE the record and own the
    # simulation-method compatibility axes.
    assert "rng_version" in section["effective_records"]["AUTO"]["fields"]
    assert "sim_method_version" in section["effective_records"]["AUTO"]["fields"]


def test_77_the_seed_domain_is_not_restated_by_this_grammar() -> None:
    section = _request_section()
    assert section["supplied_seed_domain_owner"] == "input_contract.yaml"
    for key in ("seed_min", "seed_max", "minimum", "maximum", "range"):
        assert key not in section, key
    # TYPE and PRESENCE here; admissibility there.
    assert section["field_types"]["supplied_seed"] == "F_I"
    inputs = load_contract(CONTRACT_PATH)
    from pccm_builder.sim_rng import _seed_domain

    assert _seed_domain(inputs) == (1, 2147483646)


def test_78_the_excluded_fields_are_still_excluded_and_never_encoded() -> None:
    section = _request_section()
    excluded = ["effective_seed", "auto_nonce", "run_id", "selected_confidence_level"]
    assert list(section["excluded_fields"]) == excluded
    for name in excluded:
        assert name not in section["fields"], name
        for shape in section["effective_records"].values():
            assert name not in shape["fields"], name
        for production in section["grammar"].values():
            assert name not in production, name
    assert section["analytical_fingerprint_hashed_as_a_field"] is False
    for absent in ("result_digest", "timestamp", "model_version", "contingency"):
        for production in section["grammar"].values():
            assert absent not in production, absent


def test_79_the_analytical_sections_remain_an_untouched_prefix() -> None:
    block = _raw()["request_fingerprint"]
    assert block["section_order"] == ["HEADER", "COST", "RISK", "SIM"]
    assert block["analytical_prefix"] == ["HEADER", "COST", "RISK"]
    assert block["section_order"][:3] == block["analytical_prefix"]
    assert block["extension_semantics"] == "prefix_plus_extension"
    assert block["existing_sections_modified"] is False


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
