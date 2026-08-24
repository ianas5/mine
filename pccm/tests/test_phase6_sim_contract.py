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
    assert block["tolerance"] is None


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
    assert sim.reserved_rows_h == 32
    assert sim.layout.header_row == 32
    assert sim.layout.first_iteration_row == 33
    assert sim.layout.footer_rows == 0
    assert sim.max_iterations_representable == MAX_EXCEL_ROWS - 32 == 1048544
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
    assert _raw()["sim_state"]["never_evaluated_status"] is None


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
    text = SIM_PATH.read_text(encoding="utf-8")
    body = "\n".join(
        line for line in text.splitlines() if not line.lstrip().startswith("#")
    )
    lowered = body.lower()
    assert "tolerance: null" in lowered
    assert lowered.count("tolerance") == 1, "the only tolerance mention is the null statement"
    for token in ("rel_tol", "abs_tol", "rtol", "atol", "ulp"):
        assert token not in lowered, token


def test_46_every_authority_reference_resolves() -> None:
    sim = _sim()
    assert len(sim.authority_references) == 10
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


def test_48_no_production_module_imports_the_evidence_package() -> None:
    """Tests may read retained evidence for conformance. Production must not."""
    builder = PCCM_ROOT / "builder"
    offenders = []
    for path in sorted(builder.rglob("*.py")):
        text = path.read_text(encoding="utf-8")
        if "evidence/phase6_step0" in text or "phase6_step0" in text:
            offenders.append(str(path.relative_to(PCCM_ROOT)))
    assert not offenders, f"production modules reference the evidence package: {offenders}"


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


def test_50_no_phase6_vba_or_emission_exists() -> None:
    src = PCCM_ROOT / "src" / "vba"
    names = {p.name for p in src.glob("*.bas")} | {p.name for p in src.glob("*.cls")}
    for banned in ("modSimRng.bas", "modSimEngine.bas", "modSimReport.bas",
                   "modSimContract.bas"):
        assert banned not in names, f"{banned} exists; Step 1 authorises no Phase-6 VBA"
    for path in sorted(src.glob("*.bas")):
        text = path.read_text(encoding="utf-8", errors="replace")
        assert "MRG32k3a" not in text, f"{path.name} references MRG32k3a"


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
