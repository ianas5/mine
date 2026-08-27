#!/usr/bin/env python3
"""PCCM Phase 6 Step-1 negative tests for the simulation contract.

`spec/sim_contract.yaml` is the sixth authority. Two things have to be enforced,
not merely documented:

  AUTHORITY  the contract must ENCODE what Step 0 settled, not choose it. Every
             RNG constant, jump element, Cheng literal, expression order,
             comparison operator, version ownership and digest field is checked
             against a locked constant in `sim_loader.py`. A ONE-TOKEN mutation
             must fail.

  BOUNDARY   the contract must not acquire authority it was not given. The
             admissible seed range belongs to `input_contract.yaml`, the business
             iteration minimum belongs there too, and every oracle comparison
             tolerance belongs to the evidence policy outside the contract. A
             second copy of any of them is the drift this architecture prevents.

A validator that cannot detect a one-token authority mutation is not sufficient,
so every rule below plants a defect and asserts the refusal fires.

NO VBA IS EXECUTED HERE, and nothing simulates.

Runs standalone or under pytest.
"""

from __future__ import annotations

import copy
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable

import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    SimContractError,
    StructureContractError,
    load_contract,
    load_driver_contract,
    load_sim_contract,
    load_spec,
    load_structure_contract,
    validate_sim_against,
)

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"
CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
DRIVERS_PATH = PCCM_ROOT / "spec" / "driver_contract.yaml"
STRUCTURE_PATH = PCCM_ROOT / "spec" / "structure_contract.yaml"
CALC_PATH = PCCM_ROOT / "spec" / "calc_contract.yaml"
SIM_PATH = PCCM_ROOT / "spec" / "sim_contract.yaml"

_BASE: dict[str, Any] | None = None
_STRUCTURE_BASE: dict[str, Any] | None = None


def _base() -> dict[str, Any]:
    global _BASE
    if _BASE is None:
        _BASE = yaml.safe_load(SIM_PATH.read_text(encoding="utf-8"))
    return copy.deepcopy(_BASE)


def _structure_base() -> dict[str, Any]:
    global _STRUCTURE_BASE
    if _STRUCTURE_BASE is None:
        _STRUCTURE_BASE = yaml.safe_load(STRUCTURE_PATH.read_text(encoding="utf-8"))
    return copy.deepcopy(_STRUCTURE_BASE)


def _write(data: dict[str, Any], tmp: str, name: str = "broken.yaml") -> Path:
    path = Path(tmp) / name
    path.write_text(
        yaml.safe_dump(data, sort_keys=False, allow_unicode=True), encoding="utf-8"
    )
    return path


def _rejected(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    """The mutated contract must fail at load time."""
    data = _base()
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-badsim-") as tmp:
        path = _write(data, tmp)
        try:
            load_sim_contract(path)
        except SimContractError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of SimContractError"
            ) from error
    raise AssertionError(f"{reason}: an invalid simulation contract was silently accepted")


def _rejected_cross(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    """The contract must fail CROSS validation against the owning authorities."""
    data = _base()
    with tempfile.TemporaryDirectory(prefix="pccm-badsim-") as tmp:
        sim = load_sim_contract(_write(data, tmp))
        contract_doc = yaml.safe_load(CONTRACT_PATH.read_text(encoding="utf-8"))
        mutate(contract_doc)
        broken = _write(contract_doc, tmp, "broken_input.yaml")
        try:
            validate_sim_against(
                sim,
                load_spec(SPEC_PATH),
                load_contract(broken),
                load_driver_contract(DRIVERS_PATH),
                load_structure_contract(STRUCTURE_PATH),
                yaml.safe_load(CALC_PATH.read_text(encoding="utf-8")),
            )
        except SimContractError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of SimContractError"
            ) from error
    raise AssertionError(f"{reason}: a broken authority boundary was silently accepted")


def _structure_rejected(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    data = _structure_base()
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-badstruct-") as tmp:
        path = _write(data, tmp)
        try:
            load_structure_contract(path)
        except StructureContractError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of StructureContractError"
            ) from error
    raise AssertionError(f"{reason}: an invalid structure contract was silently accepted")


# ===========================================================================
# the control that makes every other control meaningful
# ===========================================================================
def test_00_the_unmutated_contract_loads() -> None:
    """If the base contract failed, every rejection below would prove nothing."""
    with tempfile.TemporaryDirectory(prefix="pccm-goodsim-") as tmp:
        load_sim_contract(_write(_base(), tmp))


# ===========================================================================
# shape and parser boundary
# ===========================================================================
def test_01_a_missing_section_is_rejected() -> None:
    for section in ("rng", "cheng", "risk", "sim_data", "result_digest", "versions",
                    "iterations", "authority_references"):
        _rejected(lambda d, s=section: d.pop(s), f"missing section {section}")


def test_02_a_wrong_document_version_is_rejected() -> None:
    _rejected(lambda d: d.__setitem__("sim_contract_version", "2.0.0"), "wrong document version")


def test_03_a_duplicate_key_is_rejected_at_the_parser_boundary() -> None:
    text = SIM_PATH.read_text(encoding="utf-8")
    doubled = text.replace(
        "sim_contract_version: \"1.0.0\"",
        "sim_contract_version: \"1.0.0\"\nsim_contract_version: \"9.9.9\"",
        1,
    )
    with tempfile.TemporaryDirectory(prefix="pccm-dupsim-") as tmp:
        path = Path(tmp) / "dup.yaml"
        path.write_text(doubled, encoding="utf-8")
        try:
            load_sim_contract(path)
        except SimContractError as error:
            assert "duplicate key" in str(error)
            return
    raise AssertionError("a duplicate key was silently resolved last-wins")


# ===========================================================================
# B. authoritative constants
# ===========================================================================
def test_04_a_wrong_rng_constant_is_rejected() -> None:
    for key in ("m1", "m2", "a12", "a13n", "a21", "a23n"):
        _rejected(
            lambda d, k=key: d["rng"]["constants"].__setitem__(k, d["rng"]["constants"][k] + 1),
            f"rng constant {key} off by one",
        )


def test_05_a_perturbed_norm_is_rejected() -> None:
    import math

    _rejected(
        lambda d: d["rng"]["constants"].__setitem__(
            "norm", math.nextafter(d["rng"]["constants"]["norm"], math.inf)
        ),
        "norm perturbed by one ULP",
    )


def test_06_an_extra_rng_constant_is_rejected() -> None:
    _rejected(lambda d: d["rng"]["constants"].__setitem__("a11", 1), "unknown rng constant")


def test_07_a_reversed_state_order_is_rejected() -> None:
    _rejected(
        lambda d: d["rng"]["state"].__setitem__("order", list(reversed(d["rng"]["state"]["order"]))),
        "state order reversed",
    )
    _rejected(
        lambda d: d["rng"]["state"].__setitem__("orientation", "newest_first"),
        "state orientation flipped",
    )


def test_08_a_mutated_uniform_combination_is_rejected() -> None:
    _rejected(
        lambda d: d["rng"]["combination"].__setitem__(
            "rule", "if p1 < p2 then u = (p1 - p2 + m1) * norm else u = (p1 - p2) * norm"
        ),
        "combination comparison changed from <= to <",
    )
    _rejected(
        lambda d: d["rng"]["combination"].__setitem__("comparison_operator", "less_than"),
        "combination operator relabelled",
    )


def test_09_an_inclusive_uniform_endpoint_is_rejected() -> None:
    for key in ("lower_inclusive", "upper_inclusive"):
        _rejected(
            lambda d, k=key: d["rng"]["output_domain"].__setitem__(k, True),
            f"uniform endpoint {key} made inclusive",
        )


def test_10_permitting_naive_floating_modulo_is_rejected() -> None:
    _rejected(
        lambda d: d["rng"]["arithmetic"].__setitem__("naive_floating_modulo_permitted", True),
        "naive floating modulo permitted",
    )


# ===========================================================================
# C. Cheng formulation
# ===========================================================================
def test_11_log4_substituted_for_the_literal_is_rejected() -> None:
    def mutate(d):
        lines = d["cheng"]["bb"]["per_attempt"]
        d["cheng"]["bb"]["per_attempt"] = [
            line.replace("1.3862944", "log(4)") for line in lines
        ]

    _rejected(mutate, "log(4) substituted for the locked literal 1.3862944")
    _rejected(
        lambda d: d["cheng"]["bb"]["literals"].__setitem__(0, "log(4)"),
        "literal list records log(4) instead of 1.3862944",
    )


def test_12_the_log1p_formulation_is_rejected() -> None:
    def mutate(d):
        d["cheng"]["bb"]["per_attempt"] = [
            line.replace("log(u1 / (1 - u1))", "log(u1) - log1p(-u1)")
            for line in d["cheng"]["bb"]["per_attempt"]
        ]

    _rejected(mutate, "log1p formulation substituted for the locked logit form")
    _rejected(
        lambda d: d["cheng"].__setitem__("logit_form", "log(u1) - log1p(-u1)"),
        "logit_form itself replaced by the rejected alternative",
    )


def test_13_a_bb_bc_boundary_mutation_is_rejected() -> None:
    _rejected(
        lambda d: d["distributions"]["beta_pert"]["dispatch"].__setitem__(
            "rule", "min(alpha, beta) >= 1 -> BB; otherwise BC"
        ),
        "dispatch boundary moved from > to >=",
    )
    _rejected(
        lambda d: d["distributions"]["beta_pert"]["dispatch"].__setitem__(
            "equality_belongs_to", "BB"
        ),
        "equality reassigned from BC to BB",
    )
    _rejected(
        lambda d: d["distributions"]["beta_pert"]["dispatch"].__setitem__(
            "comparison_operator", "greater_than_or_equal"
        ),
        "dispatch comparison operator relabelled",
    )


def test_14_an_inverted_bc_orientation_is_rejected() -> None:
    """The silent defect: a valid Beta variate of the MIRRORED distribution."""
    _rejected(
        lambda d: d["cheng"]["bc"]["orientation"].update(
            {"a": "min(alpha0, beta0)", "b": "max(alpha0, beta0)"}
        ),
        "BC orientation inverted to match BB",
    )
    _rejected(
        lambda d: d["cheng"]["bb"]["orientation"].update(
            {"a": "max(alpha0, beta0)", "b": "min(alpha0, beta0)"}
        ),
        "BB orientation inverted to match BC",
    )


def test_15_a_reordered_cheng_expression_list_is_rejected() -> None:
    def mutate(d):
        lines = d["cheng"]["bb"]["per_attempt"]
        lines[3], lines[4] = lines[4], lines[3]

    _rejected(mutate, "BB per-attempt expression order swapped")


def test_16_a_mutated_bc_literal_is_rejected() -> None:
    for index, replacement in ((0, "1/72"), (1, "3/72"), (2, "7/9")):
        _rejected(
            lambda d, i=index, r=replacement: d["cheng"]["bc"]["literals"].__setitem__(i, r),
            f"BC literal {index} expressed as the fraction {replacement}",
        )


def test_17_a_changed_attempt_uniform_count_is_rejected() -> None:
    for value in (1, 3):
        _rejected(
            lambda d, v=value: d["cheng"].__setitem__(
                "uniforms_per_non_degenerate_proposal_attempt", v
            ),
            f"proposal attempt consuming {value} uniforms",
        )


def test_18_a_broken_source_binding_is_rejected() -> None:
    _rejected(
        lambda d: d["cheng"]["source_binding"].__setitem__("functions_sha256", "not-a-digest"),
        "malformed Cheng source binding",
    )


def test_19_treating_the_vectors_as_a_runtime_table_is_rejected() -> None:
    _rejected(
        lambda d: d["cheng"]["conformance_vectors"].__setitem__("runtime_lookup_table", True),
        "conformance vectors relabelled as a runtime lookup table",
    )


# ===========================================================================
# D. jump
# ===========================================================================
def test_20_a_single_altered_jump_element_is_rejected() -> None:
    for matrix in ("a1_p127", "a2_p127"):
        for i in range(3):
            for j in range(3):
                _rejected(
                    lambda d, m=matrix, r=i, c=j: d["jump"][m][r].__setitem__(
                        c, d["jump"][m][r][c] + 1
                    ),
                    f"{matrix}[{i}][{j}] off by one",
                )


def test_21_a_malformed_jump_matrix_is_rejected() -> None:
    _rejected(lambda d: d["jump"].__setitem__("a1_p127", [[1, 2], [3, 4], [5, 6]]),
              "jump matrix with the wrong width")
    _rejected(lambda d: d["jump"].__setitem__("a1_p127", [[1, 2, 3], [4, 5, 6]]),
              "jump matrix with the wrong height")


def test_22_a_changed_jump_exponent_or_decomposition_is_rejected() -> None:
    _rejected(lambda d: d["jump"].__setitem__("stream_spacing_exponent", 76),
              "stream spacing changed to 2^76")
    _rejected(lambda d: d["jump"].__setitem__("substream_spacing_exponent", 76),
              "substreams reintroduced")
    _rejected(lambda d: d["jump"].__setitem__("decomposition_h", 1 << 16),
              "MultModM H changed")
    _rejected(lambda d: d["jump"].__setitem__("naive_floating_matrix_product_permitted", True),
              "naive floating matrix product permitted")


# ===========================================================================
# E. seed authority
# ===========================================================================
def test_23_a_seed_range_copied_into_the_sim_contract_is_rejected() -> None:
    _rejected(
        lambda d: d["seeding"].__setitem__("seed_min", 1),
        "seed minimum duplicated into the simulation contract",
    )
    _rejected(
        lambda d: d["seeding"].__setitem__("seed_max", 2147483646),
        "seed maximum duplicated into the simulation contract",
    )
    _rejected(
        lambda d: d["seeding"].__setitem__("admissible_seed_domain", [1, 2147483646]),
        "the admissible domain restated in the simulation contract",
    )


def test_24_a_bare_seed_maximum_anywhere_else_is_rejected() -> None:
    _rejected(
        lambda d: d["statistics"].__setitem__("note", 2147483646),
        "the seed-domain maximum appearing outside the nonce cycle",
    )


def test_25_a_wrong_input_contract_seed_range_is_rejected() -> None:
    _rejected_cross(
        lambda d: d["inputs"]["random_seed"]["validation"].__setitem__("formula2", "2147483647"),
        "input-contract seed maximum widened past the accepted domain",
    )
    _rejected_cross(
        lambda d: d["inputs"]["random_seed"]["validation"].__setitem__("formula1", "0"),
        "input-contract seed minimum lowered to 0",
    )
    _rejected_cross(
        lambda d: d["inputs"]["random_seed"].__setitem__("validation", None),
        "input-contract seed validation removed",
    )
    _rejected_cross(
        lambda d: d["inputs"]["random_seed"]["validation"].__setitem__(
            "operator", "greaterThanOrEqual"
        ),
        "input-contract seed rule reduced to a minimum only",
    )


def test_26_forbidding_a_blank_seed_is_rejected() -> None:
    """Blank is not an omission. It is the AUTO request."""
    _rejected_cross(
        lambda d: d["inputs"]["random_seed"]["validation"].__setitem__("allow_blank", False),
        "blank forbidden on the Random Seed input",
    )
    _rejected_cross(
        lambda d: d["inputs"]["random_seed"].__setitem__("required", True),
        "Random Seed made mandatory",
    )


def test_27_a_mixer_or_alternate_seed_expansion_is_rejected() -> None:
    _rejected(
        lambda d: d["seeding"]["scalar_to_state"].__setitem__("mixer", "splitmix64"),
        "a seed mixer introduced",
    )
    _rejected(
        lambda d: d["seeding"]["scalar_to_state"].__setitem__("rule", "modular_expansion"),
        "the rejected D6-05 expansion selected",
    )


# ===========================================================================
# F. nonce semantics
# ===========================================================================
def test_28_a_reordered_nonce_lifecycle_is_rejected() -> None:
    def mutate(d):
        order = d["seeding"]["nonce_lifecycle"]["order"]
        i = order.index("read_current_auto_nonce")
        j = order.index("persist_auto_nonce_plus_one")
        order[i], order[j] = order[j], order[i]

    _rejected(mutate, "nonce lifecycle reordered to advance before reading")


def test_29_a_lifecycle_step_dropped_is_rejected() -> None:
    _rejected(
        lambda d: d["seeding"]["nonce_lifecycle"]["order"].remove("persist_auto_nonce_plus_one"),
        "the nonce advance step removed",
    )


def test_30_making_a_post_allocation_failure_free_is_rejected() -> None:
    _rejected(
        lambda d: d["seeding"]["nonce_lifecycle"].__setitem__(
            "failure_after_allocation_consumes_nonce", False
        ),
        "an allocated seed reused after a failure",
    )
    _rejected(
        lambda d: d["seeding"]["nonce_lifecycle"].__setitem__(
            "failure_before_allocation_consumes_nonce", True
        ),
        "a validation refusal spending a nonce",
    )


def test_31_a_mutated_exhaustion_value_is_rejected() -> None:
    _rejected(
        lambda d: d["seeding"]["nonce_lifecycle"].__setitem__("exhausted_value", 2147483645),
        "exhaustion value moved",
    )
    _rejected(
        lambda d: d["seeding"]["nonce_lifecycle"].__setitem__("on_exhaustion", "WRAP"),
        "wrapping at exhaustion",
    )
    _rejected(
        lambda d: d["seeding"]["nonce_lifecycle"].__setitem__("wrap_permitted", True),
        "wrap explicitly permitted",
    )


def test_32_stepped_multiplication_as_the_authority_is_rejected() -> None:
    _rejected(
        lambda d: d["seeding"]["auto"].__setitem__("stepped_multiplication_is_the_authority", True),
        "O(nonce) stepped multiplication declared the authority",
    )
    _rejected(
        lambda d: d["seeding"]["auto"].__setitem__("mapping_kind", "stepped_multiplication"),
        "mapping kind changed away from modular exponentiation",
    )


def test_33_claiming_cross_workbook_uniqueness_is_rejected() -> None:
    _rejected(
        lambda d: d["seeding"]["auto"].__setitem__("cross_workbook_uniqueness_claimed", True),
        "cross-workbook AUTO uniqueness claimed",
    )
    _rejected(
        lambda d: d["seeding"]["auto"].__setitem__("timestamp_derived_uniqueness_permitted", True),
        "timestamp-derived uniqueness permitted",
    )


# ===========================================================================
# G. stream assignment
# ===========================================================================
def test_34_physical_row_order_is_rejected() -> None:
    _rejected(
        lambda d: d["stream_assignment"].__setitem__("physical_row_order_permitted", True),
        "physical row order permitted for stream assignment",
    )
    _rejected(
        lambda d: d["stream_assignment"].__setitem__("policy", "physical_row_order"),
        "stream assignment policy set to physical row order",
    )
    _rejected(
        lambda d: d["accumulation"].__setitem__("physical_row_order_permitted", True),
        "physical row order permitted for accumulation",
    )


def test_35_numeric_id_sorting_is_rejected() -> None:
    _rejected(
        lambda d: d["stream_assignment"].__setitem__(
            "numeric_suffix_interpretation_permitted", True
        ),
        "numeric interpretation of the Permanent-ID suffix permitted",
    )
    _rejected(
        lambda d: d["stream_assignment"].__setitem__("permanent_id_comparison", "numeric"),
        "Permanent-ID comparison changed to numeric",
    )
    _rejected(
        lambda d: d["stream_assignment"].__setitem__(
            "permanent_id_comparison", "locale_collation"
        ),
        "Permanent-ID comparison changed to locale collation",
    )


def test_36_a_wrong_component_kind_set_or_order_is_rejected() -> None:
    def reorder(d):
        kinds = d["components"]["kinds"]
        kinds[1], kinds[2] = kinds[2], kinds[1]

    _rejected(reorder, "component kind order swapped")
    _rejected(
        lambda d: d["components"]["kinds"].append(
            {"key": "RISK_SEVERITY", "driver_kind": "risk", "per_driver": 1, "role": "severity"}
        ),
        "duplicate component kind",
    )
    _rejected(
        lambda d: d["components"]["kinds"][0].__setitem__("per_driver", 2),
        "a cost line given two sampling streams",
    )
    _rejected(
        lambda d: d["stream_assignment"].__setitem__("sort_keys", ["permanent_id", "role"]),
        "component kind dropped from the sort key",
    )


def test_37_hiding_the_accepted_stream_shift_consequence_is_rejected() -> None:
    _rejected(
        lambda d: d["stream_assignment"].__setitem__(
            "accepted_consequence", "stream_identities_are_stable_across_driver_insertion"
        ),
        "the accepted stream-shift consequence contradicted",
    )


# ===========================================================================
# H. D6-18
# ===========================================================================
def test_38_a_conditional_severity_policy_is_rejected() -> None:
    _rejected(
        lambda d: d["risk"]["severity"].__setitem__("invocation_policy", "conditional"),
        "severity invocation made conditional on occurrence",
    )
    _rejected(
        lambda d: d["risk"]["severity"].__setitem__(
            "sampler_invoked_every_risk_iteration", False
        ),
        "severity sampler no longer invoked every iteration",
    )


def test_39_the_withdrawn_advancement_wording_is_rejected() -> None:
    _rejected(
        lambda d: d["risk"]["severity"].__setitem__(
            "note", "the severity stream advances once per iteration"
        ),
        "the withdrawn 'advances once per iteration' wording reintroduced",
    )


def test_40_a_degenerate_driver_consuming_a_uniform_is_rejected() -> None:
    _rejected(
        lambda d: d["distributions"]["degenerate"].__setitem__("uniforms_consumed", 1),
        "degenerate driver consuming one uniform",
    )
    _rejected(
        lambda d: d["risk"]["severity"].__setitem__("degenerate_consumption", 1),
        "degenerate severity consuming one uniform",
    )
    _rejected(
        lambda d: d["distributions"]["degenerate"].__setitem__("stream_state_changed", True),
        "degenerate driver advancing its stream",
    )
    _rejected(
        lambda d: d["distributions"]["degenerate"].__setitem__(
            "detected_before_parameterisation", False
        ),
        "degeneracy detected after parameterisation, where 0/0 can arise",
    )


def test_41_a_non_strict_occurrence_comparison_is_rejected() -> None:
    _rejected(
        lambda d: d["risk"]["occurrence"].__setitem__(
            "comparison_operator", "less_than_or_equal"
        ),
        "occurrence comparison weakened to <=",
    )
    _rejected(
        lambda d: d["risk"]["occurrence"].__setitem__(
            "rule", "occurred = u_occurrence <= probability"
        ),
        "occurrence rule weakened to <=",
    )


def test_42_folding_probability_into_k_factors_is_rejected() -> None:
    for key in ("probability_folded_into_knom", "probability_folded_into_kpv"):
        _rejected(lambda d, k=key: d["risk"].__setitem__(k, True), f"{key} enabled")


# ===========================================================================
# I. result digest
# ===========================================================================
def test_43_a_digest_tag_mutation_is_rejected() -> None:
    _rejected(lambda d: d["result_digest"].__setitem__("stream_tag", "PCCM-FP"),
              "result stream tag collided with the input fingerprint tag")
    _rejected(lambda d: d["result_digest"].__setitem__("section_name", "RESULTS"),
              "digest section renamed")


def test_44_a_digest_field_order_mutation_is_rejected() -> None:
    def swap(d):
        fields = d["result_digest"]["record_fields"]
        fields[1], fields[2] = fields[2], fields[1]

    _rejected(swap, "nominal and PV swapped in the digest record")
    _rejected(
        lambda d: d["result_digest"]["record_fields"].remove("iteration_index"),
        "iteration index dropped from the digest record",
    )
    _rejected(
        lambda d: d["result_digest"].__setitem__("record_field_count", 4),
        "digest record field count disagreeing with the field list",
    )
    _rejected(
        lambda d: d["result_digest"].__setitem__("iteration_index_origin", 0),
        "digest iteration index rebased to 0",
    )


def test_45_sorting_samples_for_the_digest_is_rejected() -> None:
    _rejected(
        lambda d: d["result_digest"].__setitem__("samples_sorted_for_digest", True),
        "samples sorted before digesting, destroying iteration identity",
    )
    _rejected(
        lambda d: d["result_digest"].__setitem__("order_source", "sorted_ascending"),
        "digest order taken from the sorted copies",
    )


def test_46_a_digest_version_ownership_mutation_is_rejected() -> None:
    _rejected(
        lambda d: d["result_digest"].__setitem__("version_field_source", "result_digest_version"),
        "a third result-digest version introduced",
    )
    _rejected(
        lambda d: d["versions"].__setitem__("result_digest_version_source", "rng_version"),
        "digest version reassigned to RNG_VERSION",
    )


def test_47_an_approximate_digest_comparison_is_rejected() -> None:
    _rejected(lambda d: d["result_digest"].__setitem__("equality", "approximate"),
              "digest equality weakened to approximate")
    _rejected(lambda d: d["result_digest"].__setitem__("tolerance", 1e-12),
              "a tolerance attached to the digest")
    _rejected(lambda d: d["result_digest"].__setitem__("tolerance", None),
              "a NULL tolerance field, which is still the semantic in the wrong file")


# ===========================================================================
# J. request fingerprint
# ===========================================================================
def test_48_sim_before_risk_is_rejected() -> None:
    _rejected(
        lambda d: d["request_fingerprint"].__setitem__(
            "section_order", ["HEADER", "COST", "SIM", "RISK"]
        ),
        "SIM inserted before RISK, breaking the Phase-5 prefix",
    )
    _rejected(
        lambda d: d["request_fingerprint"].__setitem__(
            "section_order", ["SIM", "HEADER", "COST", "RISK"]
        ),
        "SIM placed first",
    )


def test_49_adding_selected_confidence_level_to_the_request_is_rejected() -> None:
    _rejected(
        lambda d: d["request_fingerprint"]["sim_section"]["fields"].append(
            "selected_confidence_level"
        ),
        "Selected Confidence Level added to the execution fingerprint",
    )


def test_50_adding_run_scoped_identities_to_the_request_is_rejected() -> None:
    for field in ("effective_seed", "auto_nonce", "run_id"):
        _rejected(
            lambda d, f=field: d["request_fingerprint"]["sim_section"]["fields"].append(f),
            f"{field} added to the request fingerprint",
        )
        _rejected(
            lambda d, f=field: d["request_fingerprint"]["sim_section"]["excluded_fields"].remove(f),
            f"{field} dropped from the exclusion list",
        )


def test_51_hashing_the_analytical_fingerprint_as_a_field_is_rejected() -> None:
    _rejected(
        lambda d: d["request_fingerprint"]["sim_section"].__setitem__(
            "analytical_fingerprint_hashed_as_a_field", True
        ),
        "the analytical fingerprint hashed as a SIM field instead of extended",
    )


def test_52_modifying_the_existing_sections_is_rejected() -> None:
    _rejected(
        lambda d: d["request_fingerprint"].__setitem__("existing_sections_modified", True),
        "existing HEADER/COST/RISK bytes declared modifiable",
    )
    _rejected(
        lambda d: d["request_fingerprint"].__setitem__(
            "analytical_prefix", ["HEADER", "COST", "RISK", "SIM"]
        ),
        "SIM folded into the analytical prefix",
    )


# ===========================================================================
# K. simulation state
# ===========================================================================
def test_53_a_fourth_simulation_state_is_rejected() -> None:
    _rejected(
        lambda d: d["sim_state"]["states"].append("NOT SIMULATED"),
        "a fourth simulation state added",
    )
    _rejected(
        lambda d: d["label_sets"]["sim_state"].append("UNSELECTED"),
        "the rejected UNSELECTED state added to the label set",
    )
    _rejected(
        lambda d: d["sim_state"].__setitem__("no_success_valid_status", "NOT SIMULATED"),
        "a fourth state introduced through the no-success field",
    )


def test_54_publishing_a_partial_distribution_is_rejected() -> None:
    _rejected(
        lambda d: d["sim_state"]["on_failure"].__setitem__("partial_distribution_published", True),
        "partial publication on failure",
    )
    _rejected(
        lambda d: d["sim_state"]["on_failure"].__setitem__("prior_sim_data_preserved", False),
        "a failure destroying the prior successful _SimData",
    )


def test_55_silent_phase5_recalculation_is_rejected() -> None:
    _rejected(
        lambda d: d["prerequisite"].__setitem__("silent_recalculation_permitted", True),
        "Phase 6 permitted to recalculate Phase 5 silently",
    )
    _rejected(
        lambda d: d["prerequisite"].__setitem__("phase5_analytical_state_required", "STALE"),
        "the Phase-5 prerequisite weakened",
    )


# ===========================================================================
# L. D6-08
# ===========================================================================
def test_56_a_free_ceiling_literal_that_disagrees_with_the_layout_is_rejected() -> None:
    _rejected(
        lambda d: d["iterations"]["technical_ceiling"].__setitem__(
            "max_iterations_representable", 1000000
        ),
        "a plausible-looking ceiling literal that the layout does not produce",
    )
    _rejected(
        lambda d: d["iterations"]["technical_ceiling"].__setitem__("reserved_rows_h", 30),
        "H declared as a free constant disagreeing with the layout",
    )


def test_57_a_layout_change_that_does_not_move_the_ceiling_is_rejected() -> None:
    """Adding a reserved row without recomputing the ceiling is the exact defect
    a derived constant exists to prevent."""

    def mutate(d):
        d["sim_data"]["reserved_rows"].append({"rows": [34, 34], "purpose": "smuggled row"})

    _rejected(mutate, "an extra reserved row with the ceiling left unchanged")


def test_58_a_gap_or_overlap_in_the_reserved_rows_is_rejected() -> None:
    def gap(d):
        d["sim_data"]["reserved_rows"][2]["rows"] = [4, 4]

    def overlap(d):
        d["sim_data"]["reserved_rows"][2]["rows"] = [2, 3]

    _rejected(gap, "a gap in the reserved-row tiling")
    _rejected(overlap, "an overlap in the reserved-row tiling")


def test_59_a_footer_that_silently_reduces_capacity_is_rejected() -> None:
    _rejected(
        lambda d: d["sim_data"]["iteration_records"].__setitem__("footer_rows", 1),
        "a footer row below the iteration records",
    )


def test_60_a_business_iteration_maximum_is_rejected() -> None:
    _rejected(
        lambda d: d["iterations"].__setitem__("business_maximum", 1000000),
        "a business maximum invented for iterations",
    )
    _rejected(
        lambda d: d["iterations"].__setitem__("business_minimum", 1000),
        "the business minimum duplicated out of the input contract",
    )
    _rejected(
        lambda d: d["iterations"]["technical_ceiling"].__setitem__(
            "presented_as_business_validation", True
        ),
        "the technical ceiling presented as business validation",
    )


def test_61_a_ceiling_refusal_that_consumes_a_nonce_is_rejected() -> None:
    _rejected(
        lambda d: d["iterations"]["technical_ceiling"].__setitem__("consumes_auto_nonce", True),
        "a storage-ceiling refusal consuming an AUTO nonce",
    )
    _rejected(
        lambda d: d["iterations"]["technical_ceiling"]["refusal_precedes"].remove(
            "auto_seed_allocation"
        ),
        "the refusal no longer required to precede seed allocation",
    )


def test_62_a_misplaced_iteration_header_row_is_rejected() -> None:
    _rejected(
        lambda d: d["sim_data"]["iteration_records"].__setitem__("header_row", 32),
        "iteration header row disagreeing with the reserved-row count",
    )
    _rejected(
        lambda d: d["sim_data"]["iteration_records"].__setitem__("first_iteration_row", 35),
        "a gap between the header row and the first iteration row",
    )


def test_63_a_sorted_or_extended_sim_data_table_is_rejected() -> None:
    _rejected(
        lambda d: d["sim_data"]["iteration_records"].__setitem__("sorted", True),
        "_SimData sorted, destroying iteration identity",
    )
    _rejected(
        lambda d: d["sim_data"]["iteration_records"]["columns"].append(
            {"key": "driver_samples", "column": "E", "header": "Samples", "value_type": "double"}
        ),
        "per-driver samples added to _SimData",
    )
    _rejected(
        lambda d: d["sim_data"]["excluded"].remove("per_driver_samples"),
        "the per-driver-sample exclusion dropped",
    )


def test_64_a_run_identity_field_dropped_is_rejected() -> None:
    for key in ("run_id", "effective_seed", "consumed_auto_nonce", "result_digest",
                "request_fingerprint"):
        def mutate(d, k=key):
            fields = d["sim_data"]["run_identity"]["fields"]
            index = next(i for i, f in enumerate(fields) if f["key"] == k)
            fields.pop(index)
            for offset, field in enumerate(fields):
                field["row"] = 8 + offset
            d["sim_data"]["run_identity"]["last_row"] = 8 + len(fields) - 1
            rows = d["sim_data"]["reserved_rows"]
            entry = next(e for e in rows if e["rows"][0] == 8)
            entry["rows"][1] -= 1
            for later in rows[rows.index(entry) + 1:]:
                later["rows"][0] -= 1
                later["rows"][1] -= 1
            records = d["sim_data"]["iteration_records"]
            records["header_row"] -= 1
            records["first_iteration_row"] -= 1
            ceiling = d["iterations"]["technical_ceiling"]
            ceiling["reserved_rows_h"] -= 1
            ceiling["max_iterations_representable"] += 1

        _rejected(mutate, f"run identity field {key} dropped")


# ===========================================================================
# M. versions
# ===========================================================================
def test_65_a_change_owned_by_no_version_is_rejected() -> None:
    for owner in ("rng_version", "sim_method_version"):
        _rejected(
            lambda d, o=owner: d["versions"]["bump_ownership"][o].pop(),
            f"a change removed from the {owner} bump list, leaving it ownerless",
        )


def test_66_a_change_moved_to_the_wrong_version_is_rejected() -> None:
    def mutate(d):
        bumps = d["versions"]["bump_ownership"]
        bumps["rng_version"].append(
            bumps["sim_method_version"].pop(
                bumps["sim_method_version"].index(
                    "cheng_formulation_literals_or_expression_order"
                )
            )
        )

    _rejected(mutate, "the Cheng formulation reassigned from SIM_METHOD to RNG_VERSION")


def test_67_an_invalid_version_value_is_rejected() -> None:
    for value in (0, -1, "1", 1.0, True):
        _rejected(lambda d, v=value: d["versions"].__setitem__("rng_version", v),
                  f"rng_version = {value!r}")


# ===========================================================================
# N. tolerance and boundaries
# ===========================================================================
def test_68_a_comparison_tolerance_anywhere_is_rejected() -> None:
    _rejected(
        lambda d: d["statistics"].__setitem__("comparison_tolerance", 1e-12),
        "a comparison tolerance added to statistics",
    )
    _rejected(
        lambda d: d["cheng"].__setitem__("output_tolerance", 1e-11),
        "a Cheng output tolerance added",
    )
    _rejected(
        lambda d: d["rng"].__setitem__("ulp_tolerance", 4),
        "a ULP tolerance added to the RNG section",
    )
    _rejected(
        lambda d: d["statistics"].__setitem__("note", "compare within tolerance 1e-10"),
        "a tolerance smuggled into a prose field",
    )


def test_69_a_missing_or_extra_authority_reference_is_rejected() -> None:
    _rejected(lambda d: d["authority_references"].pop(0), "an authority reference deleted")
    _rejected(
        lambda d: d["authority_references"].append(
            {"concept": "invented", "owner": "input_contract.yaml", "locator": "inputs"}
        ),
        "an unexpected authority reference added",
    )
    _rejected(
        lambda d: d["authority_references"][0].__setitem__("owner", "workbook.yaml"),
        "an authority reference redirected to the wrong owner",
    )
    _rejected(
        lambda d: d["authority_references"][0].__setitem__("locator", "inputs.does_not_exist"),
        "an authority reference locator moved",
    )


def test_70_an_unresolvable_authority_locator_fails_cross_validation() -> None:
    """A reference that points at nothing must fail RESOLUTION, not merely shape.

    The load-time set check (test_69) refuses a locator that differs from the
    accepted boundary set, so a broken locator cannot reach cross-validation
    through the YAML. It is injected directly into the parsed contract instead,
    which is the only way to exercise the resolver itself.
    """
    import dataclasses

    from pccm_builder.sim_loader import AuthorityReference

    sim = load_sim_contract(SIM_PATH)
    broken = dataclasses.replace(
        sim,
        authority_references=(
            AuthorityReference(
                concept="Random Seed admissible domain",
                owner="input_contract.yaml",
                locator="inputs.no_such_input",
            ),
        )
        + sim.authority_references[1:],
    )
    try:
        validate_sim_against(
            broken,
            load_spec(SPEC_PATH),
            load_contract(CONTRACT_PATH),
            load_driver_contract(DRIVERS_PATH),
            load_structure_contract(STRUCTURE_PATH),
            yaml.safe_load(CALC_PATH.read_text(encoding="utf-8")),
        )
    except SimContractError as error:
        assert "does not resolve" in str(error)
        return
    raise AssertionError("an unresolvable authority locator was silently accepted")


# ===========================================================================
# O. statistics and contingency
# ===========================================================================
def test_71_a_percentile_method_change_is_rejected() -> None:
    _rejected(
        lambda d: d["statistics"]["percentile"].__setitem__("method", "nearest_rank"),
        "percentile method changed away from Hyndman-Fan Type 7",
    )
    _rejected(
        lambda d: d["statistics"]["percentile"]["formula"].__setitem__("h", "n * p"),
        "the Type-7 h formula mutated",
    )
    _rejected(
        lambda d: d["statistics"]["percentile"]["formula"].__setitem__(
            "value", "x[lo] + f * (x[hi] - x[lo])"
        ),
        "convex interpolation replaced by the overflow-prone difference form",
    )


def test_72_a_population_standard_deviation_is_rejected() -> None:
    _rejected(
        lambda d: d["statistics"]["standard_deviation"].__setitem__("divisor", "n"),
        "population divisor n instead of n-1",
    )
    _rejected(
        lambda d: d["statistics"]["standard_deviation"].__setitem__(
            "naive_sum_of_squares_permitted", True
        ),
        "naive sum of squares permitted",
    )


def test_73_a_forbidden_contingency_baseline_is_rejected() -> None:
    for baseline in ("simulation_mean", "analytical_expected_total", "a_plus_emv"):
        _rejected(
            lambda d, b=baseline: d["contingency"].__setitem__("baseline", b),
            f"contingency baselined on {baseline}",
        )
    _rejected(
        lambda d: d["contingency"].__setitem__(
            "formula", "selected_px_total - simulation_mean"
        ),
        "contingency formula rebaselined on the simulation mean",
    )


def test_74_making_selected_cl_execution_relevant_is_rejected() -> None:
    for key in ("enters_simulation_execution", "enters_request_fingerprint",
                "affects_staleness"):
        _rejected(
            lambda d, k=key: d["statistics"]["selected_confidence_level"].__setitem__(k, True),
            f"Selected Confidence Level made {key}",
        )
    _rejected(
        lambda d: d["statistics"].__setitem__("p10_selectable", True),
        "P10 made selectable",
    )


def test_75_contracting_deferred_results_scope_is_rejected() -> None:
    _rejected(
        lambda d: d["results_minimum"].__setitem__(
            "annual_simulated_samples_contracted", True
        ),
        "annual simulated samples contracted in Step 1",
    )
    _rejected(
        lambda d: d["results_minimum"]["deferred"].remove("Sensitivity"),
        "Sensitivity removed from the deferred list",
    )


def test_76_a_fourth_distribution_family_is_rejected() -> None:
    _rejected(
        lambda d: d["distributions"]["families"].append("Normal"),
        "a fourth distribution family added",
    )
    _rejected(
        lambda d: d["distributions"]["uniform"].__setitem__("most_likely_used", True),
        "Most Likely read by the Uniform family",
    )
    _rejected(
        lambda d: d["distributions"]["beta_pert"].__setitem__("lambda", 6),
        "the PERT lambda changed",
    )


# ===========================================================================
# P. D6-11 - the scoped forbidden-construct schema
# ===========================================================================
def _forbidden(data: dict) -> list:
    return data["vba"]["forbidden_constructs"]


def test_77_the_global_scalar_entry_still_works_beside_the_one_grant() -> None:
    """Backward compatibility: the scalar shape is unchanged by the grants.

    Every rule still appears in the flattened list, and every rule EXCEPT the
    scoped ones is still global in exactly the way it was.
    """
    structure = load_structure_contract(STRUCTURE_PATH)
    assert "Rnd(" in structure.forbidden_constructs
    assert "MRG32k3a" in structure.forbidden_constructs
    assert "RunSimulation" in structure.forbidden_constructs
    scoped = []
    for rule in structure.forbidden_construct_rules:
        if rule.is_scoped:
            scoped.append(rule.construct)
            continue
        assert rule.forbidden_in("modAnything") is True, rule.construct
        assert rule.forbidden_in("modSimRng") is True, rule.construct
    assert scoped == ["MRG32k3a", "RunSimulation"], scoped


def test_78_a_valid_synthetic_scoped_entry_is_accepted() -> None:
    """A scoped exception works ONLY when its owner is a declared module.

    Synthetic on purpose, and it stays synthetic: this fixture exercises the
    SCHEMA on a contract of its own making, so it keeps proving the rule shape
    whatever the real contract grants. Step 6 has since made the first real
    grant - MRG32k3a to modSimRng - and that grant is asserted against the
    actual contract in `test_phase6_sim_rng_vba.py`, not here.
    """
    data = _structure_base()
    owner = data["vba"]["modules"][0]["name"]
    _forbidden(data).append({"construct": "SyntheticFutureConstruct", "allowed_in": [owner]})
    with tempfile.TemporaryDirectory(prefix="pccm-scoped-") as tmp:
        structure = load_structure_contract(_write(data, tmp))
    rule = next(
        r for r in structure.forbidden_construct_rules
        if r.construct == "SyntheticFutureConstruct"
    )
    assert rule.is_scoped is True
    assert rule.forbidden_in(owner) is False
    assert rule.forbidden_in("modSomethingElse") is True
    assert "SyntheticFutureConstruct" in structure.forbidden_constructs


def test_79_an_unknown_scoped_owner_is_rejected() -> None:
    """This is what stops Step 1 pre-authorising code that does not exist."""
    _structure_rejected(
        lambda d: _forbidden(d).append({"construct": "MRG32k3a", "allowed_in": ["modSimRng"]}),
        "MRG32k3a scoped to a module the contract does not declare",
    )
    _structure_rejected(
        lambda d: _forbidden(d).append(
            {"construct": "RunSimulation", "allowed_in": ["modSimReport"]}
        ),
        "RunSimulation scoped to a module the contract does not declare",
    )


def test_80_a_wildcard_owner_is_rejected() -> None:
    for wildcard in ("*", "**", "all", "any", "all_modules", "*.bas"):
        _structure_rejected(
            lambda d, w=wildcard: _forbidden(d).append(
                {"construct": "SyntheticFutureConstruct", "allowed_in": [w]}
            ),
            f"wildcard owner {wildcard!r}",
        )


def test_81_an_empty_allowed_in_is_rejected() -> None:
    _structure_rejected(
        lambda d: _forbidden(d).append(
            {"construct": "SyntheticFutureConstruct", "allowed_in": []}
        ),
        "an empty exception list, which reads as a grant but grants nothing",
    )


def test_82_a_duplicate_owner_is_rejected() -> None:
    def mutate(d):
        owner = d["vba"]["modules"][0]["name"]
        _forbidden(d).append(
            {"construct": "SyntheticFutureConstruct", "allowed_in": [owner, owner]}
        )

    _structure_rejected(mutate, "the same owner declared twice")


def test_83_a_malformed_or_ambiguous_entry_is_rejected() -> None:
    _structure_rejected(
        lambda d: _forbidden(d).append({"construct": "SyntheticFutureConstruct"}),
        "a mapping carrying only 'construct'",
    )
    _structure_rejected(
        lambda d: _forbidden(d).append({"allowed_in": ["modConstants"]}),
        "a mapping carrying only 'allowed_in'",
    )
    _structure_rejected(
        lambda d: _forbidden(d).append(
            {"construct": "X", "allowed_in": ["modConstants"], "note": "why"}
        ),
        "an unknown key in the scoped shape",
    )
    _structure_rejected(lambda d: _forbidden(d).append(["MRG32k3a"]),
                        "a list where a string or mapping is required")
    _structure_rejected(lambda d: _forbidden(d).append(""), "an empty construct string")
    _structure_rejected(lambda d: _forbidden(d).append(42), "a numeric entry")


def test_84_a_duplicate_construct_is_rejected() -> None:
    _structure_rejected(
        lambda d: _forbidden(d).append("MRG32k3a"),
        "the same construct declared twice, so two rules could disagree",
    )


def test_85_the_globally_forbidden_set_is_not_weakened() -> None:
    """Rnd(, Randomize, NPV and Percentile stay globally forbidden."""
    structure = load_structure_contract(STRUCTURE_PATH)
    scoped = {r.construct for r in structure.forbidden_construct_rules if r.is_scoped}
    for construct in ("Rnd(", "Randomize", "NPV", "Percentile"):
        assert construct in structure.forbidden_constructs, construct
        assert construct not in scoped, f"{construct} acquired a scoped exception"


# ===========================================================================
# Q. closed-world schema - unknown keys fail EVERYWHERE
# ===========================================================================
def _closed_paths() -> list[str]:
    from pccm_builder.sim_loader import CLOSED_KEYS

    return sorted(CLOSED_KEYS)


def _node_at(data: dict, path: str):
    """Reach the first mapping at a schema path, resolving `[]` to element 0."""
    node = data
    if path:
        for part in path.split("."):
            if part.endswith("[]"):
                node = node[part[:-2]][0]
            else:
                node = node[part]
    return node


def test_86_an_unknown_key_at_the_root_is_rejected() -> None:
    _rejected(lambda d: d.__setitem__("future_semantic", True), "unknown root key")


def test_87_an_unknown_key_in_EVERY_closed_mapping_is_rejected() -> None:
    """Systematic. One injection per declared mapping shape, no exceptions.

    Six of these were silently ACCEPTED before this correction, which meant a
    semantic could be added to the authority document and go entirely unread by
    the validator that exists to enforce it.
    """
    paths = _closed_paths()
    assert len(paths) >= 70, f"only {len(paths)} mappings are closed"
    for path in paths:
        def mutate(d, p=path):
            _node_at(d, p)["pccm_unknown_probe"] = "x"

        _rejected(mutate, f"unknown key injected at {path or 'root'!r}")


def test_88_a_mapping_at_an_undeclared_path_is_rejected() -> None:
    _rejected(
        lambda d: d.__setitem__("future_block", {"a": 1}),
        "a whole mapping at a path the schema does not describe",
    )
    _rejected(
        lambda d: d["rng"].__setitem__("future_block", {"a": 1}),
        "a nested mapping at an undeclared path",
    )


def test_89_an_extra_run_identity_field_is_rejected() -> None:
    """The layout is exact authority; it is not extensible by accident."""

    def mutate(d):
        d["sim_data"]["run_identity"]["fields"].append(
            {"key": "invented", "row": 30, "group": "snapshot",
             "label": "Invented", "value_type": "text", "initial": None}
        )
        d["sim_data"]["run_identity"]["last_row"] = 30

    _rejected(mutate, "an invented run-identity field appended on the next row")

    def reorder(d):
        fields = d["sim_data"]["run_identity"]["fields"]
        fields[0]["row"], fields[1]["row"] = fields[1]["row"], fields[0]["row"]
        fields[0], fields[1] = fields[1], fields[0]

    _rejected(reorder, "two run-identity fields swapped")

    _rejected(
        lambda d: d["sim_data"]["run_identity"]["fields"][11].__setitem__(
            "group", "derived"
        ),
        "model_version moved out of the snapshot group",
    )


def test_90_a_missing_model_version_is_rejected() -> None:
    def mutate(d):
        fields = d["sim_data"]["run_identity"]["fields"]
        index = next(i for i, f in enumerate(fields) if f["key"] == "model_version")
        fields.pop(index)
        for offset, field in enumerate(fields):
            field["row"] = 8 + offset
        d["sim_data"]["run_identity"]["last_row"] = 8 + len(fields) - 1
        rows = d["sim_data"]["reserved_rows"]
        entry = next(e for e in rows if e["rows"][0] == 8)
        entry["rows"][1] -= 1
        for later in rows[rows.index(entry) + 1:]:
            later["rows"][0] -= 1
            later["rows"][1] -= 1
        records = d["sim_data"]["iteration_records"]
        records["header_row"] -= 1
        records["first_iteration_row"] -= 1
        ceiling = d["iterations"]["technical_ceiling"]
        ceiling["reserved_rows_h"] -= 1
        ceiling["max_iterations_representable"] += 1

    _rejected(mutate, "model_version dropped from the Run Stamp")


# ===========================================================================
# R. the corrected state authority
# ===========================================================================
def test_91_a_state_rule_that_reads_the_attempt_history_is_rejected() -> None:
    _rejected(
        lambda d: d["sim_state"]["derivation"]["rules"][2].__setitem__(
            "condition", "request_fingerprint_matches_and_last_attempt_succeeded"
        ),
        "CURRENT made conditional on the attempt result",
    )
    _rejected(
        lambda d: d["sim_state"]["derivation"]["rules"][0].__setitem__(
            "condition", "prerequisites_refuse_or_last_attempt_failed"
        ),
        "INVALID made conditional on a FAILED attempt - revision 6's overlap",
    )
    _rejected(
        lambda d: d["sim_state"].__setitem__(
            "attempt_result_participates_in_derivation", True
        ),
        "the attempt axis declared part of the derivation",
    )
    _rejected(
        lambda d: d["sim_state"]["definitions"].__setitem__(
            "CURRENT", "the fingerprint matches and the last attempt succeeded"
        ),
        "the CURRENT definition reintroducing the attempt history",
    )


def test_92_a_reordered_or_incomplete_state_derivation_is_rejected() -> None:
    def swap(d):
        rules = d["sim_state"]["derivation"]["rules"]
        rules[0], rules[1] = rules[1], rules[0]
        rules[0]["order"], rules[1]["order"] = 1, 2

    _rejected(swap, "the INVALID and BLANK rules swapped, breaking totality order")
    _rejected(
        lambda d: d["sim_state"]["derivation"]["rules"].pop(1),
        "the no-snapshot rule removed, leaving a case with no state",
    )
    _rejected(
        lambda d: d["sim_state"]["derivation"].__setitem__("ordered", False),
        "the rules declared unordered, so they stop being mutually exclusive",
    )


def test_93_a_fourth_state_through_the_blank_field_is_rejected() -> None:
    _rejected(
        lambda d: d["sim_state"].__setitem__("no_success_valid_status", "NOT SIMULATED"),
        "the blank status given a label",
    )
    _rejected(
        lambda d: d["sim_state"]["derivation"]["rules"][1].__setitem__(
            "status", "NOT SIMULATED"
        ),
        "a fourth status returned by rule 2",
    )


# ===========================================================================
# S. the contribution contract
# ===========================================================================
def test_94_sampling_the_cost_line_total_instead_of_unit_cost_is_rejected() -> None:
    _rejected(
        lambda d: d["contribution"]["cost_line"].__setitem__("sampled_quantity", "total_cost"),
        "Cost Line total sampled instead of unit cost",
    )
    _rejected(
        lambda d: d["contribution"]["cost_line"].__setitem__(
            "total_cost_uncertainty_sampled", True
        ),
        "total-cost uncertainty sampled directly",
    )
    _rejected(
        lambda d: d["contribution"]["cost_line"].__setitem__(
            "quantity_inside_distribution", True
        ),
        "Quantity moved inside the distribution",
    )


def test_95_omitting_or_doubling_quantity_is_rejected() -> None:
    _rejected(
        lambda d: d["contribution"]["cost_line"].__setitem__("nominal", "unit_cost * Knom"),
        "Quantity omitted from the Cost Line contribution",
    )
    _rejected(
        lambda d: d["contribution"]["cost_line"].__setitem__(
            "nominal", "unit_cost * Quantity * Quantity * Knom"
        ),
        "Quantity applied twice",
    )
    _rejected(
        lambda d: d["contribution"]["cost_line"].__setitem__("quantity_applications", 2),
        "two declared Quantity applications",
    )
    _rejected(
        lambda d: d["contribution"]["cost_line"].__setitem__("quantity_is_deterministic", False),
        "Quantity declared stochastic",
    )


def test_96_folding_probability_into_risk_k_factors_is_rejected() -> None:
    _rejected(
        lambda d: d["contribution"]["risk"].__setitem__(
            "probability_folded_into_k_factors", True
        ),
        "Probability folded into the Risk K factors",
    )
    _rejected(
        lambda d: d["contribution"]["risk"].__setitem__(
            "nominal_when_occurred", "severity * Probability * Knom"
        ),
        "Probability multiplied into the Risk contribution",
    )


def test_97_multiplying_a_risk_by_quantity_is_rejected() -> None:
    _rejected(
        lambda d: d["contribution"]["risk"].__setitem__("quantity_applies", True),
        "Quantity applied to a Risk",
    )
    _rejected(
        lambda d: d["contribution"]["risk"].__setitem__(
            "nominal_when_occurred", "severity * Quantity * Knom"
        ),
        "Quantity multiplied into the Risk contribution",
    )


def test_98_ignoring_occurrence_is_rejected() -> None:
    _rejected(
        lambda d: d["contribution"]["risk"].__setitem__("nominal_when_not_occurred", "severity * Knom"),
        "a non-occurring Risk still contributing its severity",
    )
    _rejected(
        lambda d: d["contribution"]["risk"].__setitem__("pv_when_not_occurred", 1),
        "a non-occurring Risk contributing to PV",
    )
    _rejected(
        lambda d: d["contribution"]["risk"].__setitem__(
            "occurred", "occurrence_uniform <= Probability"
        ),
        "the occurrence comparison weakened in the contribution contract",
    )


def test_99_deriving_pv_from_nominal_is_rejected() -> None:
    _rejected(
        lambda d: d["contribution"].__setitem__("pv_derived_from_nominal", True),
        "PV discounted from the nominal total instead of computed with Kpv",
    )
    _rejected(
        lambda d: d["contribution"]["cost_line"].__setitem__(
            "pv", "unit_cost * Quantity * Knom * discount"
        ),
        "PV expressed through Knom rather than Kpv",
    )
    _rejected(
        lambda d: d["contribution"]["iteration_total"].__setitem__(
            "measures_independent", False
        ),
        "the two measures declared dependent",
    )


# ===========================================================================
# T. kernel, numerical domain, dependence, publication, cancellation
# ===========================================================================
def test_100_worksheet_or_com_access_in_the_hot_loop_is_rejected() -> None:
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
        _rejected(lambda d, f=flag: d["kernel"].__setitem__(f, True), f"kernel {flag} permitted")
    _rejected(
        lambda d: d["kernel"].__setitem__("inputs_resolved_once_before_simulation", False),
        "inputs resolved inside the loop",
    )
    _rejected(
        lambda d: d["kernel"]["resolved_before_loop"].remove("knom_per_driver"),
        "Knom no longer resolved before the loop",
    )


def test_101_a_correlated_driver_policy_is_rejected() -> None:
    _rejected(
        lambda d: d["dependence"].__setitem__("inter_driver_dependence", "correlated"),
        "inter-driver correlation introduced",
    )
    _rejected(
        lambda d: d["dependence"].__setitem__("correlation_matrix_supported", True),
        "a correlation matrix declared supported",
    )
    _rejected(
        lambda d: d["dependence"].__setitem__("copula_supported", True),
        "a copula declared supported",
    )
    _rejected(
        lambda d: d["dependence"].__setitem__("shared_or_hidden_dependence_permitted", True),
        "hidden dependence permitted",
    )


def test_102_a_hidden_positivity_or_magnitude_restriction_is_rejected() -> None:
    _rejected(
        lambda d: d["numerical_domain"].__setitem__("positivity_rule", "min > 0"),
        "a positivity rule invented for Phase 6",
    )
    _rejected(
        lambda d: d["numerical_domain"].__setitem__("magnitude_restriction", "abs(x) < 1e15"),
        "a magnitude restriction invented for Phase 6",
    )
    _rejected(
        lambda d: d["numerical_domain"].__setitem__("negative_values_legal", False),
        "negative values made illegal",
    )
    _rejected(
        lambda d: d["numerical_domain"].__setitem__("supports_crossing_zero_legal", False),
        "supports crossing zero made illegal",
    )
    _rejected(
        lambda d: d["numerical_domain"].__setitem__("narrower_than_phase5", True),
        "the Phase-6 domain declared narrower than Phase 5's",
    )


def test_103_an_unsafe_arithmetic_policy_is_rejected() -> None:
    _rejected(
        lambda d: d["numerical_domain"]["disciplines"].__setitem__(
            "accumulation", "naive_running_sum"
        ),
        "naive accumulation substituted for the safe signed sum",
    )
    _rejected(
        lambda d: d["numerical_domain"]["disciplines"].__setitem__(
            "percentile_interpolation", "difference_form"
        ),
        "the overflow-prone percentile difference form substituted",
    )
    _rejected(
        lambda d: d["numerical_domain"]["disciplines"].__setitem__(
            "driver_contribution", "naive_product"
        ),
        "naive multiplication substituted for the safe product",
    )
    _rejected(
        lambda d: d["numerical_domain"].__setitem__(
            "representable_result_refused_for_naive_intermediate_overflow", True
        ),
        "a valid representable result refused because a naive intermediate overflows",
    )
    _rejected(
        lambda d: d["numerical_domain"].__setitem__("silent_non_finite_result_permitted", True),
        "a silent non-finite result permitted",
    )


def test_104_enabling_cancellation_is_rejected() -> None:
    _rejected(
        lambda d: d["interruption"].__setitem__("user_cancellation_supported_in_phase_6", True),
        "user cancellation enabled in Phase 6",
    )


def test_105_partial_publication_or_results_recomputation_is_rejected() -> None:
    _rejected(
        lambda d: d["publication"].__setitem__("results_recomputes_monte_carlo", True),
        "Results recomputing its own Monte Carlo",
    )
    _rejected(
        lambda d: d["publication"].__setitem__(
            "partial_new_distribution_published_on_refusal_or_failure", True
        ),
        "a partial distribution published on failure",
    )
    _rejected(
        lambda d: d["publication"].__setitem__("commit_last", False),
        "commit-last abandoned",
    )
    _rejected(
        lambda d: d["publication"].__setitem__(
            "publish_only_after_simulation_and_statistics_complete", False
        ),
        "publication permitted before the statistics complete",
    )
    _rejected(
        lambda d: d["publication"].__setitem__("results_derives_from", "recomputation"),
        "Results no longer derived from _SimData",
    )
    _rejected(
        lambda d: d["command_surface"].__setitem__("user_facing_run_button_in_phase_6", True),
        "a Phase-6 user-facing Run Simulation button",
    )


# ===========================================================================
# U. the ladder and content-bound authority
# ===========================================================================
def test_106_dropping_the_full_ladder_rule_is_rejected() -> None:
    _rejected(
        lambda d: d["statistics"].__setitem__("include_all_selectable_ladder_values", False),
        "the full selectable ladder no longer retained",
    )
    _rejected(
        lambda d: d["statistics"].__setitem__("fixed_nonselectable_percentiles", []),
        "P10 dropped from the fixed headline percentiles",
    )
    _rejected(
        lambda d: d["statistics"].__setitem__(
            "selectable_ladder_locator", "config_tables.distributions"
        ),
        "the ladder locator redirected at the wrong table",
    )
    _rejected(
        lambda d: d["statistics"].__setitem__("p10_selectable", True),
        "P10 made selectable",
    )


def _rejected_cross_workbook(mutate, reason: str) -> None:
    """The contract must fail when the WORKBOOK authority's content drifts."""
    with tempfile.TemporaryDirectory(prefix="pccm-badwb-") as tmp:
        sim = load_sim_contract(SIM_PATH)
        document = yaml.safe_load(SPEC_PATH.read_text(encoding="utf-8"))
        mutate(document)
        broken = _write(document, tmp, "broken_workbook.yaml")
        try:
            validate_sim_against(
                sim,
                load_spec(broken),
                load_contract(CONTRACT_PATH),
                load_driver_contract(DRIVERS_PATH),
                load_structure_contract(STRUCTURE_PATH),
                yaml.safe_load(CALC_PATH.read_text(encoding="utf-8")),
            )
        except SimContractError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of SimContractError"
            ) from error
    raise AssertionError(f"{reason}: a broken content binding was silently accepted")


def test_107_simdata_visibility_drift_is_rejected() -> None:
    """Previously ACCEPTED: veryHidden -> hidden resolved, so the binding was false."""

    def mutate(d):
        sheet = next(s for s in d["sheets"] if s["name"] == "_SimData")
        sheet["visibility"] = "hidden"

    _rejected_cross_workbook(mutate, "_SimData downgraded from veryHidden to hidden")


def test_108_a_broken_results_placeholder_binding_is_rejected() -> None:
    def mutate(d):
        results = next(s for s in d["sheets"] if s["name"] == "Results")
        results["blocks"] = [b for b in results["blocks"] if b.get("title") != "Run Stamp"]

    _rejected_cross_workbook(mutate, "the Run Stamp section removed from the Results placeholder")

    def rename(d):
        results = next(s for s in d["sheets"] if s["name"] == "Results")
        for block in results["blocks"]:
            if block.get("title") == "Summary Statistics":
                block["title"] = "Statistics"

    _rejected_cross_workbook(rename, "the Summary Statistics section renamed")


class _Document:
    """A stand-in exposing a raw YAML document, for binding-level tests.

    `load_spec` and `load_contract` refuse several of these mutations at THEIR
    own boundary, which is correct layering and stronger protection. This
    bypasses them so the SIM contract's own content binding is what gets tested,
    rather than the upstream loader's.
    """

    def __init__(self, document):
        self.raw = document


def test_109_a_missing_model_version_authority_is_rejected() -> None:
    document = yaml.safe_load(SPEC_PATH.read_text(encoding="utf-8"))
    document["model"].pop("model_version")
    try:
        validate_sim_against(
            load_sim_contract(SIM_PATH),
            _Document(document),
            load_contract(CONTRACT_PATH),
            load_driver_contract(DRIVERS_PATH),
            load_structure_contract(STRUCTURE_PATH),
            yaml.safe_load(CALC_PATH.read_text(encoding="utf-8")),
        )
    except SimContractError:
        return
    raise AssertionError("a missing model_version authority was silently accepted")


def test_109b_the_upstream_loader_also_refuses_a_blank_model_version() -> None:
    """Layering, recorded: the workbook contract guards this at its own boundary."""
    from pccm_builder import SpecError

    with tempfile.TemporaryDirectory(prefix="pccm-badwb-") as tmp:
        document = yaml.safe_load(SPEC_PATH.read_text(encoding="utf-8"))
        document["model"]["model_version"] = "   "
        try:
            load_spec(_write(document, tmp, "broken_workbook.yaml"))
        except SpecError:
            return
    raise AssertionError("a blank model_version was silently accepted")


def test_110_a_distribution_master_disagreement_is_rejected() -> None:
    """Previously ACCEPTED: the master list could be changed outright."""

    def _retable(master, values):
        master["values"] = values
        master["data_rows"] = len(values)

    def drop(d):
        master = next(t for t in d["config_tables"] if t["key"] == "distributions")
        _retable(master, [v for v in master["values"] if v != "Beta-PERT"])

    def add(d):
        master = next(t for t in d["config_tables"] if t["key"] == "distributions")
        _retable(master, list(master["values"]) + ["Normal"])

    def duplicate(d):
        master = next(t for t in d["config_tables"] if t["key"] == "distributions")
        _retable(master, list(master["values"]) + ["Uniform"])

    _rejected_cross(drop, "an accepted family removed from the master list")
    _rejected_cross(add, "an extra family added to the master list")
    _rejected_cross(duplicate, "a duplicate family in the master list")


def test_111_reordering_the_master_list_is_ACCEPTED() -> None:
    """Membership is the binding, not presentation order.

    The user-facing order of `lstDistributions` has no reason to become the
    simulation's dispatch order, so a reordering must NOT fail - a validator that
    refused it would be enforcing something no authority says.
    """
    with tempfile.TemporaryDirectory(prefix="pccm-order-") as tmp:
        document = yaml.safe_load(CONTRACT_PATH.read_text(encoding="utf-8"))
        master = next(t for t in document["config_tables"] if t["key"] == "distributions")
        master["values"] = list(reversed(master["values"]))
        broken = _write(document, tmp, "reordered_input.yaml")
        validate_sim_against(
            load_sim_contract(SIM_PATH),
            load_spec(SPEC_PATH),
            load_contract(broken),
            load_driver_contract(DRIVERS_PATH),
            load_structure_contract(STRUCTURE_PATH),
            yaml.safe_load(CALC_PATH.read_text(encoding="utf-8")),
        )


def test_112_an_empty_or_missing_confidence_ladder_is_rejected() -> None:
    document = yaml.safe_load(CONTRACT_PATH.read_text(encoding="utf-8"))
    ladder = next(t for t in document["config_tables"] if t["key"] == "confidence_levels")
    ladder["values"] = []
    try:
        validate_sim_against(
            load_sim_contract(SIM_PATH),
            load_spec(SPEC_PATH),
            _Document(document),
            load_driver_contract(DRIVERS_PATH),
            load_structure_contract(STRUCTURE_PATH),
            yaml.safe_load(CALC_PATH.read_text(encoding="utf-8")),
        )
    except SimContractError:
        return
    raise AssertionError("an empty selectable ladder was silently accepted")


# ===========================================================================
# V. Uniform degeneracy must not read the ignored Most Likely
# ===========================================================================
def test_113_a_common_degeneracy_predicate_is_rejected() -> None:
    """The inherited-authority contradiction, closed.

    A single `a == m == b` for all three families made a degenerate Uniform
    depend on Most Likely, which accepted Phase-5 D1 ignores numerically and
    excludes from the calculation fingerprint. An ignored input would then decide
    RNG consumption and every later draw on that stream.
    """
    _rejected(
        lambda d: d["distributions"]["degenerate"]["conditions"].__setitem__(
            "uniform", "a == m == b"
        ),
        "Uniform degeneracy reading Most Likely",
    )
    _rejected(
        lambda d: d["distributions"]["degenerate"].__setitem__(
            "most_likely_read_by_uniform_degeneracy", True
        ),
        "Most Likely declared readable for Uniform degeneracy",
    )
    _rejected(
        lambda d: d["distributions"]["degenerate"]["conditions"].__setitem__(
            "triangular", "a == b"
        ),
        "the Triangular condition weakened to two-way",
    )
    _rejected(
        lambda d: d["distributions"]["degenerate"]["conditions"].__setitem__(
            "beta_pert", "a == b"
        ),
        "the Beta-PERT condition weakened to two-way",
    )
    _rejected(
        lambda d: d["distributions"]["uniform"].__setitem__(
            "most_likely_affects_degeneracy", True
        ),
        "Most Likely declared to affect Uniform degeneracy",
    )
    _rejected(
        lambda d: d["distributions"]["uniform"].__setitem__(
            "most_likely_affects_uniform_consumption", True
        ),
        "Most Likely declared to affect Uniform RNG consumption",
    )


# ===========================================================================
# W. pinned authority values
# ===========================================================================
def test_114_a_bumped_initial_version_is_rejected() -> None:
    """Step 0 settled both at 1. A bump is an authority change, not a value."""
    for key in ("rng_version", "sim_method_version"):
        _rejected(lambda d, k=key: d["versions"].__setitem__(k, 2),
                  f"versions.{key} bumped to 2")


def test_115_a_mutated_state_transition_is_rejected() -> None:
    _rejected(
        lambda d: d["rng"]["recurrence"].__setitem__("advance", "banana"),
        "the MRG state shift replaced with arbitrary text",
    )
    _rejected(
        lambda d: d["rng"]["recurrence"].__setitem__(
            "advance", "[s10, s11, s12, s20, s21, s22] <- [s11, s12, p1, s21, p2, s22]"
        ),
        "two words of the state shift transposed",
    )


def test_116_a_mutated_bc_expression_or_return_is_rejected() -> None:
    """BC was validated far more weakly than BB. An arbitrary BC return is how a
    mirrored distribution ships while every other check still passes."""
    _rejected(lambda d: d["cheng"]["bc"].__setitem__("per_driver", ["banana"]),
              "BC per_driver replaced")
    _rejected(lambda d: d["cheng"]["bc"]["per_attempt"].__setitem__(3, "banana"),
              "one BC per_attempt expression replaced")
    _rejected(lambda d: d["cheng"]["bc"].__setitem__("return", "banana"),
              "BC return replaced")
    _rejected(lambda d: d["cheng"]["bb"].__setitem__("return", "banana"),
              "BB return replaced")
    _rejected(
        lambda d: d["cheng"]["bc"].__setitem__(
            "return", "w / (b + w) when the caller's first parameter was the min, "
                      "else b / (b + w)"
        ),
        "BC return given BB's orientation - the mirrored-distribution defect",
    )
    _rejected(lambda d: d["cheng"]["bc"]["per_driver"].__setitem__(1, "beta = b"),
              "the BC beta setup mutated")


def test_117_a_mutated_evidence_binding_is_rejected() -> None:
    _rejected(
        lambda d: d["cheng"]["source_binding"].__setitem__("evidence_file", "banana"),
        "the Cheng formulation evidence path pointed elsewhere",
    )
    _rejected(
        lambda d: d["cheng"]["conformance_vectors"].__setitem__("evidence_file", "banana"),
        "the Cheng vectors evidence path pointed elsewhere",
    )
    _rejected(
        lambda d: d["cheng"]["source_binding"].__setitem__("functions_sha256", "0" * 64),
        "the Cheng source hash zeroed while keeping its shape",
    )


def test_118_a_zeroed_jump_hash_is_rejected() -> None:
    """Authoritative-looking metadata the validator ignores is worse than none."""
    for key in ("a1_p127_sha256", "a2_p127_sha256"):
        _rejected(lambda d, k=key: d["jump"].__setitem__(k, "0" * 64),
                  f"{key} zeroed while the matrix stayed correct")


def test_119_a_mutated_digest_grammar_is_rejected() -> None:
    for key in ("stream", "section", "record"):
        _rejected(lambda d, k=key: d["result_digest"]["grammar"].__setitem__(k, "banana"),
                  f"result_digest.grammar.{key} replaced")
    _rejected(
        lambda d: d["result_digest"]["grammar"].__setitem__(
            "record", "F_I(field_count) F_I(iteration_index) F_N(total_pv) F_N(total_nominal)"
        ),
        "nominal and PV transposed in the digest grammar",
    )
    _rejected(
        lambda d: d["result_digest"]["grammar"].__setitem__(
            "stream", 'F_S("PCCM-RD") F_I(rng_version) section'
        ),
        "the digest grammar reading RNG_VERSION instead of SIM_METHOD_VERSION",
    )


def test_120_a_mutated_conditioning_or_boundary_semantic_is_rejected() -> None:
    for family in ("triangular", "beta_pert"):
        _rejected(
            lambda d, f=family: d["distributions"][f].__setitem__(
                "conditioning_scale", "banana"
            ),
            f"{family} conditioning scale replaced",
        )
        _rejected(
            lambda d, f=family: d["distributions"][f].__setitem__(
                "conditioning_scale", "s = abs(b - a)"
            ),
            f"{family} conditioning scale changed to a width",
        )
    for key in ("m_equals_a", "m_equals_b"):
        _rejected(
            lambda d, k=key: d["distributions"]["triangular"]["boundary_cases"].__setitem__(
                k, "banana"
            ),
            f"triangular boundary case {key} replaced",
        )
    _rejected(
        lambda d: d["distributions"]["triangular"]["boundary_cases"].__setitem__(
            "m_equals_a", "c = 0; the lower branch is always taken"
        ),
        "the m = a boundary sending sampling down the wrong branch",
    )


def test_121_mutated_run_identity_initials_or_enums_are_rejected() -> None:
    def at(d, key):
        return next(f for f in d["sim_data"]["run_identity"]["fields"] if f["key"] == key)

    _rejected(lambda d: at(d, "next_auto_nonce").__setitem__("initial", 1),
              "next_auto_nonce seeded at 1")
    _rejected(lambda d: at(d, "last_run_id").__setitem__("initial", 1),
              "last_run_id seeded at 1")
    _rejected(lambda d: at(d, "last_attempt_result").__setitem__("initial", "FAILED"),
              "last_attempt_result seeded as FAILED")
    _rejected(lambda d: at(d, "simulation_status").__setitem__("initial", "CURRENT"),
              "a never-run workbook presenting a derived status")
    _rejected(lambda d: at(d, "run_id").__setitem__("initial", 0),
              "run_id seeded, making a never-run workbook look like a partial success")
    _rejected(lambda d: at(d, "result_digest").__setitem__("initial", "0000000000000000"),
              "a result digest seeded at build time")
    for key, wrong in (
        ("seed_mode", "attempt_result"),
        ("last_attempt_seed_mode", "attempt_result"),
        ("simulation_status", "attempt_result"),
        ("last_attempt_result", "sim_state"),
    ):
        _rejected(lambda d, k=key, w=wrong: at(d, k).__setitem__("enum", w),
                  f"{key} pointed at the {wrong} label set")
    _rejected(lambda d: at(d, "run_id").__setitem__("enum", "sim_state"),
              "a non-enum field declaring a label set")
    _rejected(lambda d: at(d, "run_id").__setitem__("label", "banana"),
              "a run-identity label replaced")


def test_122_mutated_run_identity_columns_are_rejected() -> None:
    for key, wrong in (("label_column", "A"), ("value_column", "C"), ("note_column", "G")):
        _rejected(
            lambda d, k=key, w=wrong: d["sim_data"]["run_identity"].__setitem__(k, w),
            f"run identity {key} moved",
        )


def test_123_mutated_iteration_columns_are_rejected() -> None:
    def col(d, key):
        return next(
            c for c in d["sim_data"]["iteration_records"]["columns"] if c["key"] == key
        )

    _rejected(lambda d: col(d, "total_nominal").__setitem__("column", "E"),
              "total_nominal moved to column E")
    _rejected(lambda d: col(d, "total_nominal").__setitem__("value_type", "text"),
              "total_nominal typed as text")
    _rejected(lambda d: col(d, "iteration_index").__setitem__("header", "banana"),
              "an iteration header replaced")
    _rejected(lambda d: col(d, "iteration_index").__setitem__("value_type", "double"),
              "the iteration index typed as a double")

    def swap(d):
        columns = d["sim_data"]["iteration_records"]["columns"]
        columns[1], columns[2] = columns[2], columns[1]

    _rejected(swap, "nominal and PV columns transposed")


def test_124_a_mutated_reserved_row_purpose_is_rejected() -> None:
    """The tiling IS the derivation of H, so its purposes are the audit trail."""
    _rejected(
        lambda d: d["sim_data"]["reserved_rows"][0].__setitem__("purpose", "banana"),
        "a reserved-row purpose replaced",
    )


def test_125_a_mutated_state_definition_is_rejected() -> None:
    _rejected(
        lambda d: d["sim_state"]["definitions"].__setitem__("CURRENT", "banana"),
        "a state definition replaced with arbitrary text",
    )
    _rejected(
        lambda d: d["stream_assignment"].__setitem__("index_rule", "banana"),
        "the stream index rule replaced",
    )


# ===========================================================================
# X. systematic sweeps
# ===========================================================================
def _leaf_paths(node, path=""):
    """Every scalar leaf, with `[]` for list elements - the schema's own idiom."""
    if isinstance(node, dict):
        for key, value in node.items():
            yield from _leaf_paths(value, f"{path}.{key}" if path else key)
    elif isinstance(node, list):
        for value in node:
            yield from _leaf_paths(value, f"{path}[]")
    else:
        yield path, node


def _mapping_keys(node, path=""):
    """Every (mapping path, key) pair, once per distinct shape."""
    seen = {}
    def walk(n, p=""):
        if isinstance(n, dict):
            for key, value in n.items():
                seen.setdefault((p, key), None)
                walk(value, f"{p}.{key}" if p else key)
        elif isinstance(n, list):
            for value in n:
                walk(value, f"{p}[]")
    walk(node, path)
    return sorted(seen)


def _tokens(path):
    """Split a schema path into (key, list_depth) steps.

    A segment may carry several `[]`, because a value can be a list of lists -
    `jump.a1_p127[][]` is the matrix's elements.
    """
    out = []
    for segment in (path.split(".") if path else []):
        depth = 0
        while segment.endswith("[]"):
            segment = segment[:-2]
            depth += 1
        out.append((segment, depth))
    return out


def _descend(node, steps):
    """Every node reachable by following `steps`, expanding lists on the way."""
    if not steps:
        yield node
        return
    (key, depth), rest = steps[0], steps[1:]
    target = node[key]
    frontier = [target]
    for _ in range(depth):
        frontier = [item for group in frontier for item in group]
    for item in frontier:
        yield from _descend(item, rest)


def _delete_at(data, path, key):
    """Delete `key` from every mapping reachable at `path`."""
    for node in _descend(data, _tokens(path)):
        if isinstance(node, dict):
            node.pop(key, None)


DELETION_SWEEP_EXPECTED_MINIMUM = 440
"""Distinct (mapping path, key) SHAPES, not instances.

`_delete_at` removes the key from EVERY mapping at the path, so one shape-level
mutation covers all of that shape's instances at once - a superset of deleting a
single one. The instance count is larger (647 in the independent sweep); the
shape count is what needs to be exhausted."""


def test_126_deleting_ANY_required_key_is_rejected() -> None:
    """The missing-key half of fail-loud.

    Closure on unknown keys alone was not enough: an independent sweep deleted
    647 keys and 55 were ACCEPTED, so a semantic could be REMOVED from the
    authority document and the validator still called it valid.

    Every deletion must now be refused. The only exemptions are the explicit
    CONDITIONAL and INTENTIONALLY_OPTIONAL entries, each with a written reason,
    and the test fails if either list grows silently.
    """
    from pccm_builder.sim_loader import CONDITIONAL_KEYS, INTENTIONALLY_OPTIONAL

    assert INTENTIONALLY_OPTIONAL == {}, (
        "a key was declared optional; every entry needs a written reason and a review"
    )
    assert set(CONDITIONAL_KEYS) == {("sim_data.run_identity.fields[]", "enum")}, (
        "the conditional-key list changed; each entry states when the key is "
        "required and when it is refused"
    )
    for reason in CONDITIONAL_KEYS.values():
        assert len(reason) > 40, "a conditional key needs a written reason"

    pairs = _mapping_keys(_base())
    assert len(pairs) >= DELETION_SWEEP_EXPECTED_MINIMUM, (
        f"only {len(pairs)} mapping keys were swept; the contract should not have shrunk"
    )
    exempt = set(CONDITIONAL_KEYS) | set(INTENTIONALLY_OPTIONAL)
    swept = 0
    for path, key in pairs:
        if (path, key) in exempt:
            continue
        swept += 1
        _rejected(
            lambda d, p=path, k=key: _delete_at(d, p, k),
            f"required key {key!r} deleted from {path or 'root'!r}",
        )
    assert swept >= DELETION_SWEEP_EXPECTED_MINIMUM - len(exempt)


def test_127_the_conditional_enum_key_is_required_where_it_applies() -> None:
    """The exemption is CONDITIONAL, not optional: both directions are enforced."""

    def drop(d):
        field = next(
            f for f in d["sim_data"]["run_identity"]["fields"] if f["key"] == "seed_mode"
        )
        field.pop("enum")

    _rejected(drop, "the enum owner dropped from an enum-typed run-identity field")

    def add(d):
        field = next(
            f for f in d["sim_data"]["run_identity"]["fields"] if f["key"] == "run_id"
        )
        field["enum"] = "sim_state"

    _rejected(add, "an enum owner added to a field that is not an enum")


SEMANTIC_SWEEP_EXPECTED_MINIMUM = 360
"""Distinct leaf PATHS. `_set_at` rewrites every leaf at the path."""


def _wrong_value(value):
    """A type-compatible wrong value, so the mutation tests SEMANTICS not shape."""
    if isinstance(value, bool):
        return not value
    if isinstance(value, int):
        return value + 1
    if isinstance(value, float):
        import math

        return math.nextafter(value, math.inf)
    if isinstance(value, str):
        return "banana"
    if value is None:
        return "banana"
    return None


def _set_at(data, path, new):
    """Set every leaf reachable at `path` to `new`, however deeply nested."""
    steps = _tokens(path)
    (key, depth) = steps[-1]
    for parent in _descend(data, steps[:-1]):
        if depth == 0:
            parent[key] = new
            continue
        frontier = [parent[key]]
        for _ in range(depth - 1):
            frontier = [item for group in frontier for item in group]
        for group in frontier:
            for index in range(len(group)):
                group[index] = new


def test_128_changing_ANY_settled_semantic_leaf_is_rejected() -> None:
    """The wrong-VALUE half.

    The unknown-key sweep proved shape closure but said nothing about content:
    `rng.recurrence.advance = "banana"` and a dozen more were accepted. Every
    settled leaf is now mutated to a type-compatible wrong value and must be
    refused.

    The flexible allow-list is deliberately tiny and each entry carries a reason.
    """
    from pccm_builder.sim_loader import FLEXIBLE_LEAVES

    assert set(FLEXIBLE_LEAVES) == {"dependence.authority"}, (
        "the flexible-leaf list changed; every entry must state why its content "
        "is descriptive rather than settled authority"
    )
    for reason in FLEXIBLE_LEAVES.values():
        assert len(reason) > 40, "a flexible leaf needs a written reason"

    leaves = {}
    for path, value in _leaf_paths(_base()):
        leaves.setdefault(path, value)
    assert len(leaves) >= SEMANTIC_SWEEP_EXPECTED_MINIMUM, (
        f"only {len(leaves)} leaves were swept; the contract should not have shrunk"
    )

    swept = 0
    for path, value in sorted(leaves.items()):
        if path in FLEXIBLE_LEAVES:
            continue
        swept += 1
        _rejected(
            lambda d, p=path, v=value: _set_at(d, p, _wrong_value(v)),
            f"settled leaf {path!r} changed to a type-compatible wrong value",
        )
    assert swept >= SEMANTIC_SWEEP_EXPECTED_MINIMUM - len(FLEXIBLE_LEAVES)


def test_129_the_flexible_leaf_really_is_flexible() -> None:
    """The one exemption is exercised, so the allow-list is not decorative."""
    data = _base()
    data["dependence"]["authority"] = "restated for a different reader"
    with tempfile.TemporaryDirectory(prefix="pccm-flex-") as tmp:
        load_sim_contract(_write(data, tmp))


# ===========================================================================
# Step-10A - the request-fingerprint grammar controls
#
# Before this closure the contract locked the SEMANTIC SIM fields and stopped
# there, so every mutation below produced a DIFFERENT byte stream that the
# validator accepted. Each one must now be refused.
# ===========================================================================
def _request(data: dict[str, Any]) -> dict[str, Any]:
    return data["request_fingerprint"]["sim_section"]


def test_130_a_sim_record_count_other_than_one_is_rejected() -> None:
    def mutate(data):
        _request(data)["record_count"] = 5
    _rejected(mutate, "five SIM records")


def test_131_iterations_as_a_double_field_is_rejected() -> None:
    def mutate(data):
        _request(data)["field_types"]["iterations"] = "F_N"
    _rejected(mutate, "iterations encoded as F_N")


def test_132_the_seed_mode_as_an_integer_field_is_rejected() -> None:
    def mutate(data):
        _request(data)["field_types"]["seed_mode"] = "F_I"
    _rejected(mutate, "seed_mode encoded as F_I")


def test_133_the_supplied_seed_as_a_double_field_is_rejected() -> None:
    def mutate(data):
        _request(data)["field_types"]["supplied_seed"] = "F_N"
    _rejected(mutate, "supplied_seed encoded as F_N")


def test_134_the_rng_version_as_a_double_field_is_rejected() -> None:
    def mutate(data):
        _request(data)["field_types"]["rng_version"] = "F_N"
    _rejected(mutate, "rng_version encoded as F_N")


def test_135_the_method_version_as_a_double_field_is_rejected() -> None:
    def mutate(data):
        _request(data)["field_types"]["sim_method_version"] = "F_N"
    _rejected(mutate, "sim_method_version encoded as F_N")


def test_136_auto_carrying_a_supplied_seed_is_rejected() -> None:
    def mutate(data):
        auto = _request(data)["effective_records"]["AUTO"]
        auto["fields"] = ["iterations", "seed_mode", "supplied_seed",
                          "rng_version", "sim_method_version"]
        auto["field_count"] = 5
    _rejected(mutate, "AUTO carrying a supplied seed")


def test_137_an_auto_sentinel_seed_is_rejected() -> None:
    """The F_I(0) placeholder, spelled into the grammar."""
    def mutate(data):
        grammar = _request(data)["grammar"]
        grammar["auto_record"] = (
            'F_I(5) F_I(iterations) F_S("AUTO") F_I(0) F_I(rng_version) '
            'F_I(sim_method_version)'
        )
    _rejected(mutate, "an AUTO record with a zero seed sentinel")


def test_137a_declaring_the_auto_seed_as_zero_rather_than_absent_is_rejected() -> None:
    def mutate(data):
        _request(data)["auto_supplied_seed_representation"] = "zero"
    _rejected(mutate, "AUTO seed declared as zero")


def test_138_fixed_omitting_the_supplied_seed_is_rejected() -> None:
    def mutate(data):
        fixed = _request(data)["effective_records"]["FIXED"]
        fixed["fields"] = ["iterations", "seed_mode", "rng_version", "sim_method_version"]
        fixed["field_count"] = 4
    _rejected(mutate, "FIXED without its supplied seed")


def test_139_moving_the_fixed_seed_after_the_versions_is_rejected() -> None:
    def mutate(data):
        _request(data)["effective_records"]["FIXED"]["fields"] = [
            "iterations", "seed_mode", "rng_version", "sim_method_version", "supplied_seed"]
    _rejected(mutate, "the FIXED seed moved after the version fields")


def test_140_five_one_field_records_are_rejected() -> None:
    def mutate(data):
        section = _request(data)
        section["record_count"] = 5
        section["grammar"]["section"] = 'F_S("SIM") F_I(5) sim_record*'
        section["grammar"]["auto_record"] = 'F_I(1) F_I(iterations)'
    _rejected(mutate, "the SIM section as five one-field records")


def test_141_encoding_the_field_names_is_rejected() -> None:
    def mutate(data):
        section = _request(data)
        section["encoded_field_names"] = True
        section["grammar"]["auto_record"] = (
            'F_I(8) F_S("iterations") F_I(iterations) F_S("seed_mode") F_S("AUTO") '
            'F_S("rng_version") F_I(rng_version) F_S("sim_method_version") '
            'F_I(sim_method_version)'
        )
    _rejected(mutate, "field names encoded into the record")


def test_142_the_selected_confidence_level_entering_sim_is_rejected() -> None:
    def mutate(data):
        section = _request(data)
        section["fields"] = list(section["fields"]) + ["selected_confidence_level"]
    _rejected(mutate, "selected_confidence_level in the SIM record")


def test_142a_the_selected_confidence_level_in_the_grammar_is_rejected() -> None:
    def mutate(data):
        _request(data)["grammar"]["auto_record"] = (
            'F_I(5) F_I(iterations) F_S("AUTO") F_S(selected_confidence_level) '
            'F_I(rng_version) F_I(sim_method_version)'
        )
    _rejected(mutate, "selected_confidence_level encoded into the grammar")


def test_143_the_effective_seed_entering_sim_is_rejected() -> None:
    def mutate(data):
        _request(data)["grammar"]["auto_record"] = (
            'F_I(5) F_I(iterations) F_S("AUTO") F_I(effective_seed) F_I(rng_version) '
            'F_I(sim_method_version)'
        )
    _rejected(mutate, "effective_seed encoded into the SIM record")


def test_144_the_auto_nonce_entering_sim_is_rejected() -> None:
    def mutate(data):
        _request(data)["grammar"]["auto_record"] = (
            'F_I(5) F_I(iterations) F_S("AUTO") F_I(auto_nonce) F_I(rng_version) '
            'F_I(sim_method_version)'
        )
    _rejected(mutate, "auto_nonce encoded into the SIM record")


def test_145_the_run_id_entering_sim_is_rejected() -> None:
    def mutate(data):
        _request(data)["grammar"]["fixed_record"] = (
            'F_I(6) F_I(iterations) F_S("FIXED") F_I(supplied_seed) F_I(run_id) '
            'F_I(rng_version) F_I(sim_method_version)'
        )
    _rejected(mutate, "run_id encoded into the SIM record")


def test_146_hashing_the_analytical_fingerprint_as_a_field_is_rejected() -> None:
    def mutate(data):
        _request(data)["analytical_fingerprint_hashed_as_a_field"] = True
    _rejected(mutate, "the analytical fingerprint hashed as a field")


def test_147_repeating_the_stream_tag_inside_the_extension_is_rejected() -> None:
    def mutate(data):
        _request(data)["grammar"]["section"] = (
            'F_S("PCCM-FP") F_S("SIM") F_I(1) sim_record'
        )
    _rejected(mutate, "PCCM-FP repeated inside the SIM extension")


def test_147a_repeating_the_stream_version_inside_the_extension_is_rejected() -> None:
    def mutate(data):
        _request(data)["grammar"]["section"] = 'F_S("SIM") F_I(FP_VERSION) sim_record'
    _rejected(mutate, "FP_VERSION repeated inside the SIM extension")


def test_147b_declaring_the_extension_to_carry_its_own_tag_is_rejected() -> None:
    def mutate(data):
        _request(data)["stream_tag_repeated_in_extension"] = True
    _rejected(mutate, "the extension declaring its own stream tag")


def test_148_moving_the_analytical_prefix_order_is_rejected() -> None:
    def mutate(data):
        data["request_fingerprint"]["section_order"] = ["COST", "HEADER", "RISK", "SIM"]
    _rejected(mutate, "the HEADER/COST/RISK prefix reordered")


def test_148a_putting_sim_before_the_analytical_prefix_is_rejected() -> None:
    def mutate(data):
        data["request_fingerprint"]["section_order"] = ["SIM", "HEADER", "COST", "RISK"]
    _rejected(mutate, "SIM placed before the analytical prefix")


def test_149_restating_the_seed_domain_in_this_grammar_is_rejected() -> None:
    def mutate(data):
        _request(data)["supplied_seed_domain_owner"] = "sim_contract.yaml"
    _rejected(mutate, "the seed domain claimed by the simulation contract")


def test_150_a_changed_golden_request_digest_is_detected() -> None:
    """A corpus vector cannot drift silently: the literals are pinned in
    `tests/test_phase6_request_fingerprint.py` and recomputed here."""
    from pccm_builder import load_calc_contract
    from pccm_builder.sim_cases import request_fingerprint, request_sim_section

    sim = load_sim_contract(SIM_PATH)
    calc = load_calc_contract(CALC_PATH)
    assert request_sim_section(sim, 1000, "AUTO") == "S3:SIMI1:1I1:4I4:1000S4:AUTOI1:1I1:1"
    assert request_fingerprint(sim, calc, 1000, "AUTO") == "5EAB16E15C2ECE24"
    assert request_fingerprint(sim, calc, 1000, "FIXED", 1) == "599C95E7274759B9"
    assert request_fingerprint(sim, calc, 1000, "FIXED", 2147483646) == "0010FB954CC94B53"
    assert request_fingerprint(sim, calc, 1001, "AUTO") == "4777C8BC35F0FFEF"
    # And a one-token grammar change moves the answer, which is the whole point.
    assert request_fingerprint(sim, calc, 1000, "AUTO") != request_fingerprint(
        sim, calc, 1000, "FIXED", 1)


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
