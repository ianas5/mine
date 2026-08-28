#!/usr/bin/env python3
"""PCCM Phase 6 Step-11A - the publication / Results / orchestration authority.

WHAT THIS FILE PROVES
--------------------------------------------------------------------------------
CONTRACT AND STAGE-A AUTHORITY, on Linux, now: the two-bank publication design,
the transaction order, what survives each class of failure, the persisted
summary and contingency ladders, the settled public read-accessor names, the one
Phase-5 preparation bridge, the reporting-only status of Selected CL, and the
materialised empty Results / `_SimData` shell.

NO VBA EXISTS FOR ANY OF IT. `modSimReport` is not authorised, `PCCM_RunSimulation`
is not implemented, and nothing here may be read as "a simulation published a
result". This is authority and shell only.

Runs standalone or under pytest.
"""

from __future__ import annotations

import copy
import json
import re
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable

import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    SimContractError,
    load_contract,
    load_driver_contract,
    load_sim_contract,
    load_spec,
    load_structure_contract,
    validate_sim_against,
)
from pccm_builder.sim_loader import (  # noqa: E402
    LOCKED_BANK_LABELS,
    LOCKED_CANDIDATE_TARGET,
    LOCKED_FINAL_COMMIT_FIELDS,
    LOCKED_FINAL_COMMIT_RANGE,
    LOCKED_READ_ACCESSORS,
    LOCKED_RESULTS_FORBIDDEN_FUNCTIONS,
    LOCKED_TRANSACTION_ORDER,
    MAX_EXCEL_ROWS,
)

SPEC = PCCM_ROOT / "spec"
SIM_PATH = SPEC / "sim_contract.yaml"
WORKBOOK_PATH = SPEC / "workbook.yaml"
CASES_JSON = PCCM_ROOT / "build" / "phase6_cases.json"
STAGE_A = PCCM_ROOT / "build" / "PCCM_stageA.xlsx"

_CACHE: dict[str, Any] = {}


def _sim():
    if "sim" not in _CACHE:
        _CACHE["sim"] = load_sim_contract(SIM_PATH)
    return _CACHE["sim"]


def _raw() -> dict:
    return _sim().raw


def _spec():
    if "spec" not in _CACHE:
        _CACHE["spec"] = load_spec(WORKBOOK_PATH)
    return _CACHE["spec"]


def _shell() -> dict:
    return _spec().phase6_shell


def _cases() -> dict[str, dict]:
    if "cases" not in _CACHE:
        corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
        _CACHE["cases"] = {c["id"]: c for g in corpus["groups"] for c in g["cases"]}
    return _CACHE["cases"]


def _results_shell() -> dict:
    return _shell()["results"]


def _all_results_formulas() -> list[str]:
    results = _results_shell()
    out = [f["formula"] for f in results["run_stamp"]["fields"]]
    for metric in results["summary"]["metrics"]:
        out.extend([metric["nominal"], metric["pv"]])
    selected = results["selected"]
    out.extend([selected["confidence_level_formula"], selected["quantile_nominal"],
                selected["quantile_pv"], selected["contingency_nominal"],
                selected["contingency_pv"]])
    return out


# ===========================================================================
# A. Two banks, one active
# ===========================================================================
def test_01_publication_has_exactly_two_banks() -> None:
    banks = _raw()["publication"]["banks"]
    assert tuple(banks["labels"]) == LOCKED_BANK_LABELS == ("A", "B")
    assert banks["count"] == 2
    assert banks["third_bank_permitted"] is False
    assert banks["initial_active_bank"] is None


def test_02_a_candidate_never_writes_to_the_active_bank() -> None:
    banks = _raw()["publication"]["banks"]
    assert banks["candidate_writes_to_active_bank"] is False
    assert banks["inactive_bank_is_published"] is False
    assert banks["inactive_bank_is_staging_storage"] is True
    assert dict(banks["candidate_target"]) == LOCKED_CANDIDATE_TARGET
    assert banks["candidate_target"][""] == "A", "the first success targets A"
    assert banks["candidate_target"]["A"] == "B"
    assert banks["candidate_target"]["B"] == "A"
    # No temporary sheet, no duplicate workbook, no million-row rollback.
    assert banks["temporary_worksheet_required"] is False
    assert banks["duplicate_workbook_required"] is False
    assert _raw()["publication"]["transaction"]["million_row_rollback_required"] is False


def test_03_the_active_bank_switch_is_the_final_commit() -> None:
    transaction = _raw()["publication"]["transaction"]
    assert tuple(transaction["order"]) == LOCKED_TRANSACTION_ORDER
    order = list(transaction["order"])
    assert order[-1] == "final_commit_shared_block_including_active_bank"
    assert order.index("verify_inactive_bank_against_staged_package") == len(order) - 2
    assert transaction["final_commit_range"] == LOCKED_FINAL_COMMIT_RANGE == "D22:D30"
    assert tuple(transaction["final_commit_fields"]) == LOCKED_FINAL_COMMIT_FIELDS
    assert transaction["final_commit_fields"][-1] == "active_bank"
    assert transaction["final_commit_is_one_write"] is True
    assert transaction["prior_final_commit_block_captured_before_write"] is True
    assert transaction["final_commit_failure_restores_prior_block"] is True


def test_04_the_commit_block_is_contiguous_and_shared() -> None:
    """D22:D30 must be exactly the shared rows, in row order, ending at active_bank."""
    fields = {f["key"]: f for f in _raw()["sim_data"]["run_identity"]["fields"]}
    rows = [fields[k]["row"] for k in LOCKED_FINAL_COMMIT_FIELDS]
    assert rows == list(range(22, 31)), rows
    column = _raw()["sim_data"]["run_identity"]["value_column"]
    assert LOCKED_FINAL_COMMIT_RANGE == f"{column}{rows[0]}:{column}{rows[-1]}"
    for key in LOCKED_FINAL_COMMIT_FIELDS:
        assert fields[key]["group"] in ("counter", "attempt", "derived", "control"), key
        assert fields[key]["group"] != "snapshot", f"{key} is banked, not shared"


def test_05_the_run_id_is_allocated_by_the_commit_and_by_nothing_else() -> None:
    allocation = _raw()["publication"]["run_id_allocation"]
    assert allocation["candidate_value"] == "last_run_id + 1"
    assert allocation["held_locally_until_commit"] is True
    assert allocation["allocated_by"] == "successful_final_commit"
    assert allocation["headroom_checked_before_auto_allocation"] is True
    order = list(_raw()["publication"]["transaction"]["order"])
    assert order.index("validate_pre_allocation_prerequisites") < order.index(
        "allocate_auto_nonce_when_auto")


def test_06_what_survives_each_failure_is_stated() -> None:
    failure = _raw()["publication"]["failure_semantics"]
    before = failure["refusal_before_auto_allocation"]
    assert before["next_auto_nonce_advanced"] is False
    assert before["active_bank_changed"] is False
    assert before["successful_banks_changed"] is False
    after = failure["refusal_or_failure_after_auto_allocation"]
    # NOT rolled back. A consumed AUTO sequence is consumed.
    assert after["next_auto_nonce_advanced"] is True
    assert after["active_bank_changed"] is False
    assert after["successful_banks_changed"] is False
    inactive = failure["inactive_bank_write_failure"]
    assert inactive["active_bank_changed"] is False
    assert inactive["prior_publication_remains_authoritative"] is True
    assert inactive["corrupted_candidate_has_semantic_standing"] is False
    commit = failure["final_commit_failure"]
    assert commit["prior_block_restored"] is True
    assert commit["active_bank_changed"] is False


# ===========================================================================
# B. The row ceiling does not move
# ===========================================================================
def test_07_the_second_bank_consumes_columns_not_rows() -> None:
    layout = _sim().layout
    assert layout.reserved_row_count == 33, "H moved"
    assert layout.first_iteration_row == 34
    assert layout.max_iterations_representable == MAX_EXCEL_ROWS - 33 == 1048543
    records = _raw()["sim_data"]["iteration_records"]
    assert records["header_row"] == 33
    banks = records["banks"]
    assert set(banks["A"].values()) == {"B", "C", "D"}
    assert set(banks["B"].values()) == {"F", "G", "H"}
    assert not set(banks["A"].values()) & set(banks["B"].values())
    # Bank A IS the accepted single-bank layout.
    assert {c["key"]: c["column"] for c in records["columns"]} == banks["A"]
    assert _raw()["publication"]["banks"]["row_axis_shared_by_both_banks"] is True


def test_08_the_active_bank_row_took_the_spacer_not_a_new_row() -> None:
    reserved = _raw()["sim_data"]["reserved_rows"]
    spans = [tuple(entry["rows"]) for entry in reserved]
    assert spans[-1] == (33, 33)
    assert (30, 30) not in spans, "the spacer is spent, not duplicated"
    assert any(entry["rows"] == [8, 30] for entry in reserved)
    covered = sum(last - first + 1 for first, last in spans)
    assert covered == 33


# ===========================================================================
# C. The persisted statistics and contingency ladder
# ===========================================================================
def test_09_the_full_step9_summary_is_persisted() -> None:
    summary = _raw()["sim_data"]["summary_statistics"]
    assert summary["source"] == "modSimStats"
    assert summary["recomputed_from_worksheet_data"] is False
    keys = [m["key"] for m in summary["metrics"]]
    assert keys[:3] == ["mean", "sample_standard_deviation", "minimum"]
    assert keys[-2:] == ["maximum", "deterministic_base_a"]
    assert [k for k in keys if k.startswith("quantile_")] == [
        f"quantile_{i}" for i in range(1, 12)]
    assert len(keys) == 16
    assert summary["bank_value_columns"] == {"A": {"nominal": "K", "pv": "L"},
                                             "B": {"nominal": "M", "pv": "N"}}


def test_10_no_rung_label_is_copied_into_the_simulation_contract() -> None:
    """The ladder belongs to input_contract.yaml. A rung is a projected KEY here."""
    for metric in _raw()["sim_data"]["summary_statistics"]["metrics"]:
        if metric["key"].startswith("quantile_"):
            assert metric["label"] is None, metric["key"]
    for rung in _raw()["sim_data"]["contingency_ladder"]["rungs"]:
        assert rung["label"] is None, rung["key"]


def test_11_the_whole_contingency_ladder_is_precomputed() -> None:
    ladder = _raw()["sim_data"]["contingency_ladder"]
    assert ladder["source"] == "SimStatsContingency"
    assert ladder["baseline"] == "deterministic_base_estimate_a"
    assert ladder["worksheet_subtraction_permitted"] is False
    assert ladder["computed_for_whole_ladder_before_commit"] is True
    assert ladder["all_values_representable_required_before_commit"] is True
    assert ladder["fixed_rung_persisted_though_not_selectable"] is True
    assert [r["key"] for r in ladder["rungs"]] == [f"quantile_{i}" for i in range(1, 12)]
    assert len(ladder["rungs"]) == 11
    assert ladder["bank_value_columns"] == {"A": {"nominal": "Q", "pv": "R"},
                                            "B": {"nominal": "S", "pv": "T"}}


def test_12_every_persisted_block_stays_above_the_iteration_header() -> None:
    header = _raw()["sim_data"]["iteration_records"]["header_row"]
    for key in ("summary_statistics", "contingency_ladder"):
        block = _raw()["sim_data"][key]
        assert block["last_row"] < header, key


# ===========================================================================
# D. Results is presentation
# ===========================================================================
def test_13_results_is_not_part_of_the_transaction() -> None:
    presentation = _raw()["results_minimum"]["presentation"]
    assert presentation["written_by_the_run"] is False
    assert presentation["materialised_by_stage_a"] is True
    assert _raw()["publication"]["transaction"]["results_is_a_written_transaction"] is False
    for flag in ("computes_statistics", "computes_contingency", "recomputes_quantiles",
                 "contingency_by_subtraction_on_results", "reads_a_fixed_bank"):
        assert presentation[flag] is False, flag


def test_14_every_results_formula_is_a_lookup() -> None:
    for formula in _all_results_formulas():
        upper = formula.upper()
        for banned in LOCKED_RESULTS_FORBIDDEN_FUNCTIONS:
            assert f"{banned}(" not in upper, (banned, formula[:60])
        # No arithmetic outside a string literal: no subtraction, no division.
        outside = re.sub(r'"[^"]*"', "", formula)
        for operator in ("-", "/", "*", "+", "^"):
            assert operator not in outside, (operator, formula[:80])


def test_15_every_banked_formula_reads_the_active_bank_selector() -> None:
    identity = _raw()["sim_data"]["run_identity"]
    active_row = next(f["row"] for f in identity["fields"] if f["key"] == "active_bank")
    selector = f"_SimData!${identity['value_column']}${active_row}"
    results = _results_shell()
    banked = [f["formula"] for f in results["run_stamp"]["fields"]
              if f["key"] not in ("simulation_status", "status_evaluated_at")]
    for metric in results["summary"]["metrics"]:
        banked.extend([metric["nominal"], metric["pv"]])
    for formula in banked:
        assert selector in formula, formula[:70]
        assert '="A"' in formula, "a formula does not branch on the bank label"
    # The derived status is SHARED, so it is read directly and not banked.
    for key in ("simulation_status", "status_evaluated_at"):
        formula = next(f["formula"] for f in results["run_stamp"]["fields"]
                       if f["key"] == key)
        assert selector not in formula, key


def test_16_the_selected_rows_look_up_and_never_compute() -> None:
    selected = _results_shell()["selected"]
    source = _raw()["selected_confidence_level"]["source"]
    assert source in selected["confidence_level_formula"]
    for key in ("quantile_nominal", "quantile_pv", "contingency_nominal",
                "contingency_pv"):
        formula = selected[key]
        assert "MATCH(" in formula.upper(), key
        assert "INDEX(" in formula.upper(), key
        assert source in formula, key
        assert "COUNTIF(" in formula.upper(), f"{key} does not guard an unknown selector"
    # THE FIXED RUNG IS NOT SELECTABLE: the guard range starts one row below the
    # ladder's own first row.
    summary = _raw()["sim_data"]["summary_statistics"]
    fixed_row = next(m["row"] for m in summary["metrics"] if m["key"] == "quantile_1")
    assert f"$J${fixed_row + 1}:" in selected["quantile_nominal"]
    assert f"$J${fixed_row}:" not in selected["quantile_nominal"].split("COUNTIF")[1][:20]


def test_17_selected_confidence_level_is_reporting_only() -> None:
    selector = _raw()["selected_confidence_level"]
    assert selector["source"] == "inpSelectedConfidenceLevel"
    for flag in ("participates_in_request_fingerprint",
                 "participates_in_execution_validity",
                 "participates_in_auto_allocation",
                 "participates_in_state_derivation",
                 "change_requires_rerun",
                 "invalid_selector_invalidates_simulation",
                 "unselected_state_introduced"):
        assert selector[flag] is False, flag
    assert selector["invalid_selector_blanks_selected_reporting_rows"] is True
    # And it is not a SIM field.
    sim_fields = _raw()["request_fingerprint"]["sim_section"]["fields"]
    assert "selected_confidence_level" not in sim_fields
    excluded = _raw()["request_fingerprint"]["sim_section"]["excluded_fields"]
    assert "selected_confidence_level" in excluded


# ===========================================================================
# E. The bridge and the settled surface
# ===========================================================================
def test_18_there_is_one_phase5_preparation_bridge() -> None:
    bridge = _raw()["phase5_bridge"]
    assert bridge["owner_module"] == "modCalcReport"
    assert bridge["procedure"] == "CalcPrepareSimulationInputs"
    assert not bridge["procedure"].startswith("PCCM_")
    assert bridge["is_automation_endpoint"] is False
    assert bridge["name_prefix_pccm"] is False
    assert bridge["reuses_private_preparation"] == "PrepareCurrentCalculation"
    assert bridge["requires_phase5_status"] == "CURRENT"
    assert bridge["writes_to_calc_sheet"] is False
    assert bridge["updates_phase5_status_or_attempt_metadata"] is False
    assert bridge["duplicates_factor_mathematics"] is False
    assert bridge["zero_driver_model_succeeds"] is True
    assert bridge["analytical_fingerprint_is_current_not_stored"] is True
    assert tuple(bridge["returns"]) == (
        "drivers", "driver_count", "analytical_fingerprint",
        "deterministic_base_nominal", "deterministic_base_pv", "applied_timeline",
        "decimal_separator")


def test_19_the_public_read_accessors_are_settled_and_exact() -> None:
    surface = _raw()["command_surface"]
    assert surface["read_accessor_names_settled"] is True
    assert tuple(surface["read_accessors"]) == LOCKED_READ_ACCESSORS
    assert surface["automation_endpoint"] == "PCCM_RunSimulation"
    assert surface["automation_endpoint"] not in surface["read_accessors"]
    assert tuple(surface["read_accessor_semantics"]) == LOCKED_READ_ACCESSORS
    stored = surface["read_accessor_semantics"]["PCCM_SimulationRequestFingerprint"]
    current = surface["read_accessor_semantics"]["PCCM_CurrentSimulationRequestFingerprint"]
    assert "stored" in stored and "ACTIVE bank" in stored
    assert "recomputed" in current
    assert stored != current, "the two fingerprint accessors must stay different things"
    assert surface["run_id_public_accessor_required_in_phase_6"] is False
    assert surface["effective_seed_public_accessor_required_in_phase_6"] is False
    assert surface["user_facing_run_button_in_phase_6"] is False


def test_20_the_implementation_arrived_in_step_11_and_only_there() -> None:
    """Step 11A closed the authority with NO implementation; Step 11 built it.

    The test moved with that authorisation rather than being deleted: it still
    says exactly which module may carry the endpoint, the bridge and the bank
    machinery, and no other module may.
    """
    from pccm_builder.vba_source import load_modules

    src = PCCM_ROOT / "src" / "vba"
    names = {p.name for p in src.glob("*.bas")}
    assert "modSimReport.bas" in names
    for module in load_modules([src]):
        if module.name == "modSimReport":
            continue
        for banned in ("PCCM_RunSimulation", "SIM_BANK_", "SIM_ACTIVE_BANK_ROW",
                       "SIM_FINAL_COMMIT_RANGE", "SIM_SUMMARY_", "SIM_CONTINGENCY_"):
            assert banned not in module.code, f"{module.name} carries {banned}"
        if module.name != "modCalcReport":
            assert "CalcPrepareSimulationInputs" not in module.code, module.name
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    declared = {m.name for m in structure.vba_modules}
    assert "modSimReport" in declared
    # AND NOTHING BEYOND IT. The Phase-6 registry is exactly the generated
    # projection plus the six hand-written modules, so a Step-12 module cannot
    # be declared ahead of the step that authorises it.
    assert {n for n in declared if n.startswith("modSim")} == {
        "modSimContract", "modSimRng", "modSimSample", "modSimEngine",
        "modSimStats", "modSimFingerprint", "modSimNonce", "modSimReport",
    }, sorted(n for n in declared if n.startswith("modSim"))


# ===========================================================================
# F. The materialised Stage-A shell
# ===========================================================================
def _workbook():
    if "wb" not in _CACHE:
        from openpyxl import load_workbook

        _CACHE["wb"] = load_workbook(STAGE_A)
    return _CACHE["wb"]


def test_21_the_results_shell_is_materialised_where_the_manifest_says() -> None:
    sheet = _workbook()["Results"]
    results = _results_shell()
    label = results["label_column"]
    for section in results["sections"]:
        assert sheet[f"{label}{section['row']}"].value == section["title"]
    for field_ in results["run_stamp"]["fields"]:
        assert sheet[f"{label}{field_['row']}"].value == field_["label"]
        assert sheet[f"{results['nominal_column']}{field_['row']}"].value == field_["formula"]
    for metric in results["summary"]["metrics"]:
        assert sheet[f"{label}{metric['row']}"].value == metric["label"]
        assert sheet[f"{results['nominal_column']}{metric['row']}"].value == metric["nominal"]
        assert sheet[f"{results['pv_column']}{metric['row']}"].value == metric["pv"]
    selected = results["selected"]
    assert sheet[f"{label}{selected['quantile_row']}"].value == "Selected Px"
    assert sheet[f"{label}{selected['contingency_row']}"].value == "Contingency"
    for deferred in results["deferred"]:
        assert sheet[f"{label}{deferred['row']}"].value == deferred["title"]


def test_22_the_sim_data_shell_carries_labels_and_no_data() -> None:
    sheet = _workbook()["_SimData"]
    identity = _raw()["sim_data"]["run_identity"]
    label = identity["label_column"]
    for field_ in identity["fields"]:
        assert sheet[f"{label}{field_['row']}"].value == field_["label"], field_["key"]
        if field_["group"] == "snapshot":
            for column in identity["bank_value_columns"].values():
                assert sheet[f"{column}{field_['row']}"].value is None, field_["key"]
    assert sheet[f"{identity['value_column']}21"].value == 0
    assert sheet[f"{identity['value_column']}22"].value == 0
    assert sheet[f"{identity['value_column']}23"].value == "NONE"
    for row in (24, 25, 26, 27, 28, 29, 30):
        assert sheet[f"{identity['value_column']}{row}"].value is None, row


def test_23_both_iteration_banks_have_headers_and_no_rows() -> None:
    sheet = _workbook()["_SimData"]
    records = _raw()["sim_data"]["iteration_records"]
    for bank, columns in records["banks"].items():
        for column in columns.values():
            assert sheet[f"{column}{records['header_row']}"].value is not None, (bank, column)
    for row in range(records["first_iteration_row"], records["first_iteration_row"] + 25):
        for column in "BCDEFGH":
            assert sheet[f"{column}{row}"].value is None, (row, column)


def test_24_the_persisted_blocks_carry_their_labels_and_no_values() -> None:
    sheet = _workbook()["_SimData"]
    from pccm_builder.sim_oracle import resolve_percentile_ladder

    ladder = resolve_percentile_ladder(_sim(), load_contract(SPEC / "input_contract.yaml"))
    for key, entries_key in (("summary_statistics", "metrics"),
                             ("contingency_ladder", "rungs")):
        block = _raw()["sim_data"][key]
        for entry in block[entries_key]:
            label = entry["label"]
            if label is None:
                label = ladder.ordered[int(entry["key"].split("_")[1]) - 1]
            assert sheet[f"{block['label_column']}{entry['row']}"].value == label, entry["key"]
            for columns in block["bank_value_columns"].values():
                for column in columns.values():
                    assert sheet[f"{column}{entry['row']}"].value is None, entry["key"]


def test_25_the_workbook_carries_no_simulation_output() -> None:
    sheet = _workbook()["_SimData"]
    text = " ".join(str(sheet[f"{c}{r}"].value)
                    for r in range(1, 40) for c in "BCDEFGHJKLMNPQRST"
                    if sheet[f"{c}{r}"].value is not None)
    for banned in ("PCCM-FP", "PCCM-RD", "CURRENT", "STALE", "SUCCESS"):
        assert banned not in text, banned
    assert "NONE" in text, "the initial attempt state must be materialised"


# ===========================================================================
# G. The corpus
# ===========================================================================
def test_26_the_publication_group_is_present_and_exact() -> None:
    corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
    group = [g for g in corpus["groups"] if g["group"] == "J_publication"]
    assert len(group) == 1
    ids = [c["id"] for c in group[0]["cases"]]
    assert len(ids) == len(set(ids)) == 16
    for case in group[0]["cases"]:
        assert case["comparison"] == "EXACT", case["id"]
        assert "tolerance" not in json.dumps(case)
    for required in ("publication.initial", "publication.first_success_targets_a",
                     "publication.second_success_targets_b",
                     "publication.refusal_before_auto_allocation",
                     "publication.failure_after_auto_allocation",
                     "publication.inactive_bank_write_failure",
                     "publication.final_commit_failure_restores_the_block",
                     "publication.run_id_exhaustion_refuses_first",
                     "publication.selected_cl_is_reporting_only"):
        assert required in ids, required


def test_27_the_publication_vectors_say_what_they_must() -> None:
    cases = _cases()
    initial = cases["publication.initial"]["expected_exact"]["publication_state"]
    assert initial["active_bank"] is None
    assert initial["next_auto_nonce"] == 0 and initial["last_run_id"] == 0
    assert initial["last_attempt_result"] == "NONE"

    first = cases["publication.first_success_targets_a"]["expected_exact"]["after"]
    assert first["active_bank"] == "A" and first["last_run_id"] == 1
    second = cases["publication.second_success_targets_b"]["expected_exact"]["after"]
    assert second["active_bank"] == "B"
    assert second["bank_a_request_fingerprint"] == "FP-1", "bank A must survive"

    refusal = cases["publication.refusal_before_auto_allocation"]
    assert refusal["expected_exact"]["after"]["next_auto_nonce"] == (
        refusal["inputs"]["before"]["next_auto_nonce"])
    after = cases["publication.failure_after_auto_allocation"]
    assert after["expected_exact"]["after"]["next_auto_nonce"] == (
        after["inputs"]["before"]["next_auto_nonce"] + 1)
    assert after["expected_exact"]["after"]["active_bank"] == "B"

    commit = cases["publication.final_commit_failure_restores_the_block"]
    assert commit["expected_exact"]["after"]["active_bank"] == "B"
    assert commit["expected_exact"]["after"]["last_run_id"] == 2

    exhausted = cases["publication.run_id_exhaustion_refuses_first"]
    assert exhausted["expected_exact"]["after"]["next_auto_nonce"] == (
        exhausted["inputs"]["before"]["next_auto_nonce"])

    for identifier, expected in (("publication.status.invalid", "INVALID"),
                                 ("publication.status.blank_no_success", None),
                                 ("publication.status.current", "CURRENT"),
                                 ("publication.status.stale", "STALE")):
        assert cases[identifier]["expected_exact"]["simulation_status"] == expected

    selector = cases["publication.selected_cl_is_reporting_only"]["expected_exact"]
    assert selector["request_fingerprint_changed"] is False
    assert selector["simulation_status_changed"] is False
    assert selector["selected_lookup_changed"] is True
    invalid = cases["publication.invalid_selected_cl_blanks_the_lookup"]["expected_exact"]
    assert invalid["simulation_status"] == "CURRENT"
    assert invalid["selected_quantile_displayed"] is None


# ===========================================================================
# H. Mutation controls
# ===========================================================================
def _base() -> dict:
    return copy.deepcopy(yaml.safe_load(SIM_PATH.read_text(encoding="utf-8")))


def _write(data: dict, tmp: str, name: str = "broken.yaml") -> Path:
    path = Path(tmp) / name
    path.write_text(yaml.safe_dump(data, sort_keys=False, allow_unicode=True),
                    encoding="utf-8")
    return path


def _rejected(mutate: Callable[[dict], None], reason: str) -> None:
    data = _base()
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-pub-") as tmp:
        try:
            load_sim_contract(_write(data, tmp))
        except SimContractError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__}") from error
    raise AssertionError(f"{reason}: silently accepted")


def _shell_rejected(mutate: Callable[[dict], None], reason: str) -> None:
    """A cross-validation control: the WORKBOOK shell is damaged, not the contract."""
    spec = load_spec(WORKBOOK_PATH)
    shell = copy.deepcopy(spec.phase6_shell)
    mutate(shell)
    damaged = type(spec)(
        manifest_version=spec.manifest_version, model=spec.model,
        workbook=spec.workbook, presentation=spec.presentation, sheets=spec.sheets,
        source_path=spec.source_path, phase6_shell=shell)
    try:
        validate_sim_against(
            _sim(), damaged, load_contract(SPEC / "input_contract.yaml"),
            load_driver_contract(SPEC / "driver_contract.yaml"),
            load_structure_contract(SPEC / "structure_contract.yaml"),
            yaml.safe_load((SPEC / "calc_contract.yaml").read_text(encoding="utf-8")))
    except SimContractError:
        return
    raise AssertionError(f"{reason}: silently accepted")


def test_28_a_candidate_that_writes_to_the_active_bank_is_rejected() -> None:
    _rejected(lambda d: d["publication"]["banks"].update(
        {"candidate_writes_to_active_bank": True}), "candidate writes to the active bank")


def test_29_a_third_bank_is_rejected() -> None:
    def mutate(d):
        d["publication"]["banks"]["labels"] = ["A", "B", "C"]
        d["publication"]["banks"]["count"] = 3
    _rejected(mutate, "a third bank")
    _rejected(lambda d: d["publication"]["banks"].update({"third_bank_permitted": True}),
              "a third bank permitted")


def test_30_a_first_success_that_targets_b_is_rejected() -> None:
    _rejected(lambda d: d["publication"]["banks"]["candidate_target"].update({"": "B"}),
              "the first success targets B")


def test_31_switching_the_active_bank_before_verification_is_rejected() -> None:
    def mutate(d):
        order = d["publication"]["transaction"]["order"]
        order[-1], order[-2] = order[-2], order[-1]
    _rejected(mutate, "the switch precedes verification")


def test_32_the_active_bank_switch_not_being_last_is_rejected() -> None:
    def mutate(d):
        fields = d["publication"]["transaction"]["final_commit_fields"]
        fields.remove("active_bank")
        fields.insert(0, "active_bank")
    _rejected(mutate, "active_bank is not the last committed field")


def test_33_last_run_id_outside_the_final_commit_is_rejected() -> None:
    def mutate(d):
        d["publication"]["transaction"]["final_commit_fields"].remove("last_run_id")
    _rejected(mutate, "last_run_id leaves the final commit")


def test_34_allocating_the_nonce_before_the_prerequisites_is_rejected() -> None:
    def mutate(d):
        order = d["publication"]["transaction"]["order"]
        i = order.index("allocate_auto_nonce_when_auto")
        j = order.index("validate_pre_allocation_prerequisites")
        order[i], order[j] = order[j], order[i]
    _rejected(mutate, "the nonce is consumed before the prerequisites")
    _rejected(lambda d: d["publication"]["run_id_allocation"].update(
        {"headroom_checked_before_auto_allocation": False}),
        "run-id headroom is not a pre-allocation prerequisite")


def test_35_rolling_the_auto_nonce_back_is_rejected() -> None:
    _rejected(lambda d: d["publication"]["failure_semantics"][
        "refusal_or_failure_after_auto_allocation"].update(
        {"next_auto_nonce_advanced": False}), "the consumed nonce is rolled back")


def test_36_treating_the_inactive_bank_as_published_is_rejected() -> None:
    _rejected(lambda d: d["publication"]["banks"].update(
        {"inactive_bank_is_published": True}), "the inactive bank is published")
    _rejected(lambda d: d["publication"]["failure_semantics"][
        "inactive_bank_write_failure"].update(
        {"corrupted_candidate_has_semantic_standing": True}),
        "a corrupted candidate has standing")


def test_37_results_computing_a_statistic_is_rejected() -> None:
    _rejected(lambda d: d["results_minimum"]["presentation"].update(
        {"computes_statistics": True}), "Results computes a statistic")
    _shell_rejected(
        lambda s: s["results"]["summary"]["metrics"][0].update(
            {"nominal": "=AVERAGE(_SimData!$C$34:$C$1048576)"}),
        "a Results formula averages the iteration rows")


def test_38_results_subtracting_a_contingency_is_rejected() -> None:
    _rejected(lambda d: d["results_minimum"]["presentation"].update(
        {"contingency_by_subtraction_on_results": True}), "Results subtracts")
    _shell_rejected(
        lambda s: s["results"]["selected"].update(
            {"contingency_nominal": "=_SimData!$K$21-_SimData!$K$23"}),
        "a Results formula subtracts the base from the quantile")


def test_39_results_reading_a_fixed_bank_is_rejected() -> None:
    _rejected(lambda d: d["results_minimum"]["presentation"].update(
        {"reads_a_fixed_bank": True}), "Results pins one bank")

    def mutate(shell):
        for field_ in shell["results"]["run_stamp"]["fields"]:
            field_["formula"] = field_["formula"].replace(
                '=IF(_SimData!$D$30="","",IF(_SimData!$D$30="A",', "=IF(TRUE,")
    _shell_rejected(mutate, "every Run Stamp formula pins bank A")


def test_40_a_results_field_the_contract_does_not_have_is_rejected() -> None:
    _shell_rejected(
        lambda s: s["results"]["run_stamp"]["fields"][0].update({"key": "invented"}),
        "Results presents a field the contract does not have")
    _shell_rejected(
        lambda s: s["results"]["run_stamp"]["fields"].pop(),
        "Results omits a contracted Run Stamp field")
    _shell_rejected(
        lambda s: s["results"]["summary"]["metrics"].pop(0),
        "Results omits a persisted metric")


def test_41_a_results_label_that_disagrees_with_the_contract_is_rejected() -> None:
    _shell_rejected(
        lambda s: s["results"]["run_stamp"]["fields"][1].update({"label": "Identifier"}),
        "a Run Stamp label drifts from the contract")


def test_42_a_selected_row_that_stops_looking_up_is_rejected() -> None:
    _shell_rejected(
        lambda s: s["results"]["selected"].update({"quantile_nominal": '=_SimData!$K$21'}),
        "the selected quantile stops looking the selector up")


def test_43_selected_cl_entering_the_request_fingerprint_is_rejected() -> None:
    _rejected(lambda d: d["selected_confidence_level"].update(
        {"participates_in_request_fingerprint": True}), "Selected CL enters the request")
    _rejected(lambda d: d["selected_confidence_level"].update(
        {"invalid_selector_invalidates_simulation": True}),
        "an invalid selector invalidates the simulation")
    _rejected(lambda d: d["selected_confidence_level"].update(
        {"participates_in_state_derivation": True}), "Selected CL decides staleness")


def test_44_dropping_the_persisted_summary_is_rejected() -> None:
    def mutate(d):
        d["sim_data"]["summary_statistics"]["metrics"] = [
            m for m in d["sim_data"]["summary_statistics"]["metrics"]
            if m["key"] != "mean"]
    _rejected(mutate, "the persisted summary loses a metric")
    _rejected(lambda d: d["sim_data"]["summary_statistics"].update(
        {"recomputed_from_worksheet_data": True}), "the summary is recomputed")


def test_45_a_partial_contingency_ladder_is_rejected() -> None:
    def mutate(d):
        d["sim_data"]["contingency_ladder"]["rungs"] = (
            d["sim_data"]["contingency_ladder"]["rungs"][:1])
    _rejected(mutate, "only the fixed rung is persisted")
    _rejected(lambda d: d["sim_data"]["contingency_ladder"].update(
        {"computed_for_whole_ladder_before_commit": False}),
        "the ladder is computed lazily")
    _rejected(lambda d: d["sim_data"]["contingency_ladder"].update(
        {"worksheet_subtraction_permitted": True}), "worksheet subtraction is permitted")


def test_46_rebuilding_the_phase5_inputs_independently_is_rejected() -> None:
    _rejected(lambda d: d["phase5_bridge"].update({"duplicates_factor_mathematics": True}),
              "the bridge rebuilds the factors")
    _rejected(lambda d: d["phase5_bridge"].update({"requires_phase5_status": "STALE"}),
              "the bridge accepts a stale Phase 5")
    _rejected(lambda d: d["phase5_bridge"].update(
        {"analytical_fingerprint_is_current_not_stored": False}),
        "the bridge may hand back a stored fingerprint")


def test_47_an_invented_public_name_is_rejected() -> None:
    _rejected(lambda d: d["command_surface"]["read_accessors"].append("PCCM_SimulationMean"),
              "an invented public accessor")
    _rejected(lambda d: d["command_surface"].update({"read_accessor_names_settled": False}),
              "the names become unsettled again")

    def mutate(d):
        d["command_surface"]["read_accessor_semantics"][
            "PCCM_SimulationRequestFingerprint"] = (
            d["command_surface"]["read_accessor_semantics"][
                "PCCM_CurrentSimulationRequestFingerprint"])
    _rejected(mutate, "the stored and current fingerprint accessors become one thing")


def test_48_moving_the_row_ceiling_is_rejected() -> None:
    def mutate(d):
        rows = d["sim_data"]["reserved_rows"]
        rows.append({"rows": [34, 34], "purpose": "second bank header"})
        d["sim_data"]["iteration_records"]["header_row"] = 34
        d["sim_data"]["iteration_records"]["first_iteration_row"] = 35
    _rejected(mutate, "the second bank buys itself a row")

    def share(d):
        d["sim_data"]["iteration_records"]["banks"]["B"]["iteration_index"] = "B"
    _rejected(share, "the two banks share a column")


def test_49_a_bank_that_overwrites_the_shared_state_is_rejected() -> None:
    _rejected(lambda d: d["sim_data"]["run_identity"]["bank_value_columns"].update({"B": "D"}),
              "bank B lands on the shared column")
    _rejected(lambda d: d["sim_data"]["run_identity"].update({"note_column": "F"}),
              "the note column collides with bank B")


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
