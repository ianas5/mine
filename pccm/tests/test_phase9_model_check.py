#!/usr/bin/env python3
"""PCCM Phase 9 Step 2 - the Model Check surface, its aggregation and its limits.

WHAT THESE CONTROLS ACTUALLY DO. They EVALUATE THE FORMULAS THE BUILD WROTE.
A control that re-implemented the aggregation in Python would prove that two
implementations agree, which is not the claim worth making: the sheet is the
implementation. So the scenarios below feed readings to a small Excel expression
evaluator and read the summary and the register out of the sheet's own cells.
Mutate a formula and the answer moves; that is what makes the mutation block at
the end mean anything.

NOTHING HERE IS A WINDOWS CLAIM. Excel has not run. What is settled here is what
Linux can settle: the geometry, the ordering, the counts, the precedence, the
overflow arithmetic, the duplicate rule, the advisory boundary, the threshold's
single owner, and the fact that no cell-called path reaches a writer. Whether
Excel re-evaluates an anchored cell, and whether a UDF in a cell behaves as the
five accepted adapters do, is for the Windows acceptance and is not asserted.
"""

from __future__ import annotations

import copy
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

import pytest
import yaml

PCCM_ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = PCCM_ROOT.parent
SPEC = PCCM_ROOT / "spec"
SRC = PCCM_ROOT / "src" / "vba"
MANIFEST = SPEC / "workbook.yaml"
CONTRACT_PATH = SPEC / "input_contract.yaml"

sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from excel_formula_eval import Evaluator, ExcelError, NA  # noqa: E402
from pccm_builder import (  # noqa: E402
    ContractError,
    ModelCheckPlan,
    load_contract,
    load_spec,
)
from pccm_builder.phase9_model_check import (  # noqa: E402
    ADVISORY_INPUT_KEY,
    INSPECTION_FILENAME,
    build_phase9_inspection,
    validate_phase9_inspection,
)

_CACHE: dict[str, Any] = {}


# ===========================================================================
# THE SURFACE UNDER TEST
# ===========================================================================
def _plan(manifest: dict[str, Any] | None = None,
          contract_path: Path | None = None) -> ModelCheckPlan:
    """The plan, from the shipped manifest or from a mutated copy of it."""
    if manifest is None and contract_path is None:
        if "plan" not in _CACHE:
            spec = load_spec(MANIFEST)
            shell = spec.phase6_shell
            _CACHE["plan"] = ModelCheckPlan(spec, load_contract(CONTRACT_PATH),
                                            shell["results"], shell["sensitivity"])
        return _CACHE["plan"]
    directory = Path(tempfile.mkdtemp(prefix="pccm-p9-"))
    path = directory / "workbook.yaml"
    if manifest is None:
        manifest = yaml.safe_load(MANIFEST.read_text(encoding="utf-8"))
    path.write_text(yaml.safe_dump(manifest, sort_keys=False), encoding="utf-8")
    spec = load_spec(path)
    contract = load_contract(contract_path or CONTRACT_PATH)
    shell = spec.phase6_shell
    return ModelCheckPlan(spec, contract, shell["results"], shell["sensitivity"])


def _manifest() -> dict[str, Any]:
    return yaml.safe_load(MANIFEST.read_text(encoding="utf-8"))


def _cells(plan: ModelCheckPlan) -> dict[str, Any]:
    """Every Model Check cell that holds a value or a formula, from the plan."""
    cells: dict[str, Any] = {}
    for entry in plan.summary["rows"]:
        cells[f"{plan.value_column}{entry['row']}"] = plan.summary_formula(str(entry["key"]))
    for entry in plan.readings["rows"]:
        cells[f"{plan.value_column}{entry['row']}"] = plan.reading_formula(entry)
    for position, row in enumerate(plan.register_rows(), start=1):
        for column in plan.register["columns"]:
            cells[f"{column['column']}{row}"] = plan.register_formula(
                position, str(column["key"]))
    for row, entry in plan.candidate_formulas().items():
        for key, value in entry.items():
            cells[f"{plan.candidate_column(key)}{row}"] = value
    return cells


def _built_tree() -> Path:
    if "built" not in _CACHE:
        target = Path(tempfile.mkdtemp(prefix="pccm-p9-build-"))
        result = subprocess.run(
            [sys.executable, str(PCCM_ROOT / "builder" / "build_stage_a.py"),
             "--out", str(target / "PCCM_stageA.xlsx"), "--quiet"],
            capture_output=True, text=True,
        )
        assert result.returncode == 0, result.stderr[-3000:]
        _CACHE["built"] = target
    return _CACHE["built"]


def _projection() -> dict[str, Any]:
    return json.loads((_built_tree() / INSPECTION_FILENAME).read_text(encoding="utf-8"))


# ===========================================================================
# THE SCENARIO HARNESS
# ===========================================================================
# A scenario says what each OWNER answered. Everything else - which rows fire,
# what is counted, what the summary says, what the register shows - comes out of
# the sheet's own formulas. There is no second aggregation in this file.

_BLANK_READINGS: dict[str, Any] = {
    "calculation_state": "NOT CALCULATED",
    "simulation_state": "",
    "annual_distribution_state": "NOT PRODUCED",
    "annual_profile_state": "NOT PRODUCED",
    "annual_profile_px": "",
    "annual_year_count": 0,
    "structural_report": "",
    "calculation_refusal_detail": "",
    "calculation_refusal_subject": "",
    "calculation_attempt_result": "",
    "calculation_attempt_detail": "",
    # READ, NOT TYPED. A harness that spelled the number would be exactly the
    # second owner test_04 exists to refuse - and it caught this file first.
    "requested_iterations": None,
    "simulation_publication": "",
    "published_run_id": "",
    "published_iterations_run": "",
    "simulation_status_last_evaluated": "",
    "sensitivity_availability": "No simulation has been published.",
}


def _external_key(plan: ModelCheckPlan, entry: dict[str, Any]) -> str | None:
    """Where a reading's value has to be injected, read off its own formula.

    NOT TYPED. The mirror addresses belong to the accepted Results and
    Sensitivity blocks; a harness that spelled them would go stale exactly the
    way P7-4's hand-written address did.
    """
    kind = str(entry["kind"])
    if kind == "procedure":
        return f"{entry['procedure']}()"
    if kind == "defined_name":
        return str(entry["defined_name"])
    if kind == "derived":
        return None
    formula = plan.reading_formula(entry)
    import re

    match = re.search(r"([A-Za-z_][\w ]*)!\$?([A-Z]{1,3})\$?(\d+)", formula)
    assert match is not None, formula
    return f"{match.group(1)}!{match.group(2)}{match.group(3)}"


def _workbook_default_iterations() -> int:
    """What a new workbook starts with, from the input contract that owns it."""
    return int(load_contract(CONTRACT_PATH).inputs[ADVISORY_INPUT_KEY].default)


def _evaluate(plan: ModelCheckPlan, readings: dict[str, Any] | None = None,
              cells: dict[str, Any] | None = None) -> dict[str, Any]:
    values = dict(_BLANK_READINGS)
    values["requested_iterations"] = _workbook_default_iterations()
    values.update(readings or {})
    unknown = set(values) - {str(e["key"]) for e in plan.readings["rows"]}
    assert not unknown, f"the scenario supplies readings nobody publishes: {sorted(unknown)}"

    externals: dict[str, Any] = {}
    for entry in plan.readings["rows"]:
        key = _external_key(plan, entry)
        if key is not None:
            externals[key] = values[str(entry["key"])]
    evaluator = Evaluator(cells if cells is not None else _cells(plan), externals)

    summary = {str(entry["key"]): evaluator.value(
        f"{plan.value_column}{entry['row']}") for entry in plan.summary["rows"]}
    register = []
    for row in plan.register_rows():
        cell_values = {str(column["key"]): evaluator.value(f"{column['column']}{row}")
                       for column in plan.register["columns"]}
        register.append(cell_values)
    return {"summary": summary, "register": register,
            "shown": [r for r in register if not isinstance(r["check_id"], ExcelError)],
            "unused": [r for r in register if isinstance(r["check_id"], ExcelError)]}


def _ids(result: dict[str, Any]) -> list[str]:
    return [str(row["check_id"]) for row in result["shown"]]


def _counts(result: dict[str, Any]) -> tuple[int, int]:
    summary = result["summary"]
    return int(summary["error_count"]), int(summary["warning_count"])


def _faults(count: int, key: str = "CHK_ID_PATTERN") -> str:
    """A structural report shaped exactly the way the owner shapes one."""
    return "".join(f"  [{key}] Cost Lines row {n} has a malformed permanent ID.\r\n"
                   for n in range(1, count + 1))


# ===========================================================================
# A. THE THRESHOLD HAS ONE OWNER
# ===========================================================================
def test_01_the_input_contract_owns_the_recommendation() -> None:
    contract = load_contract(CONTRACT_PATH)
    spec = contract.inputs[ADVISORY_INPUT_KEY]
    assert spec.recommended_iterations == contract.recommendation_for(ADVISORY_INPUT_KEY)
    # AND IT CHANGED NO BOUND. 1000 is still the locked hard minimum and there is
    # still no maximum.
    assert spec.validation["formula1"] == "1000"
    assert spec.validation["operator"] == "greaterThanOrEqual"
    assert "less" not in str(spec.validation.get("operator", "")).lower()


def test_02_removing_the_recommendation_fails_loudly() -> None:
    """SILENT REMOVAL IS THE FAILURE THIS FIELD EXISTS TO PREVENT. A Model Check
    whose threshold quietly vanished would simply stop advising, and nothing on
    the sheet would say so. So it fails the BUILD instead - no fallback to the
    default, no fallback to a literal, no None reaching a formula."""
    raw = yaml.safe_load(CONTRACT_PATH.read_text(encoding="utf-8"))
    del raw["inputs"][ADVISORY_INPUT_KEY]["recommended_iterations"]
    directory = Path(tempfile.mkdtemp(prefix="pccm-p9-contract-"))
    path = directory / "input_contract.yaml"
    path.write_text(yaml.safe_dump(raw, sort_keys=False), encoding="utf-8")
    contract = load_contract(path)
    with pytest.raises(ContractError, match="recommended_iterations"):
        contract.recommendation_for(ADVISORY_INPUT_KEY)
    with pytest.raises(ContractError, match="recommended_iterations"):
        _plan(contract_path=path)


def test_03_a_malformed_recommendation_fails_the_load() -> None:
    for bad in ("ten thousand", 0, -1, True, 999):
        raw = yaml.safe_load(CONTRACT_PATH.read_text(encoding="utf-8"))
        raw["inputs"][ADVISORY_INPUT_KEY]["recommended_iterations"] = bad
        directory = Path(tempfile.mkdtemp(prefix="pccm-p9-bad-"))
        path = directory / "input_contract.yaml"
        path.write_text(yaml.safe_dump(raw, sort_keys=False), encoding="utf-8")
        with pytest.raises(ContractError):
            load_contract(path)


def test_04_no_implementation_carries_the_threshold_as_a_literal() -> None:
    """ONE OWNER MEANS ONE PLACE THE NUMBER IS WRITTEN. A literal anywhere else -
    a formula, a builder, a projection, a test - is a second owner, and the two
    agree until somebody changes one."""
    plan = _plan()
    literal = str(plan.recommended_iterations)
    manifest_block = yaml.safe_dump(_manifest()["phase9_shell"], sort_keys=False)
    assert literal not in manifest_block, (
        "the Model Check manifest block carries the threshold as a literal")
    sources = [PCCM_ROOT / "builder" / "pccm_builder" / "phase9_model_check.py",
               PCCM_ROOT / "builder" / "pccm_builder" / "workbook_builder.py",
               Path(__file__)]
    for source in sources:
        text = "\n".join(line for line in source.read_text(encoding="utf-8").splitlines()
                         if not line.lstrip().startswith("#"))
        assert literal not in text, f"{source.name} types the threshold"


def test_05_the_advisory_never_reads_the_prompt_prose() -> None:
    """Advisory logic that parsed the recommendation out of the prompt sentence
    would break on a comma, a rewording or a translation, and would fail
    silently. The sentence is quoted nowhere here on purpose: quoting it would
    put the number in this file, which is the second owner test_04 refuses."""
    contract = load_contract(CONTRACT_PATH)
    prompt = contract.inputs[ADVISORY_INPUT_KEY].validation["prompt"]
    assert "recommended" in prompt.lower()
    plan = _plan()
    for entry in plan.readings["rows"]:
        assert "prompt" not in plan.reading_formula(entry).lower()
    source = (PCCM_ROOT / "builder" / "pccm_builder" / "phase9_model_check.py").read_text(
        encoding="utf-8")
    for banned in (".validation[", "validation.get(\"prompt", "prompt"):
        code = "\n".join(line for line in source.splitlines()
                         if not line.lstrip().startswith("#"))
        assert banned not in code, f"the Model Check plan reads {banned!r}"


def test_06_the_default_is_not_the_recommendation() -> None:
    """They coincide today and are different concepts. A Model Check that read
    `default` would change its advice the day a new workbook's starting value
    changed."""
    source = (PCCM_ROOT / "builder" / "pccm_builder" / "phase9_model_check.py").read_text(
        encoding="utf-8")
    code = "\n".join(line for line in source.splitlines()
                     if not line.lstrip().startswith("#"))
    assert ".default" not in code
    assert "recommendation_for" in code


# ===========================================================================
# B. GEOMETRY - AND THE BUILT SHEET IS THE PLAN
# ===========================================================================
def test_07_the_built_sheet_is_exactly_what_the_plan_says() -> None:
    """ONE LAYOUT AUTHORITY. Every other control in this file works from the
    plan; this is the control that earns the right to, by proving the workbook
    the build produced holds precisely those cells and no others."""
    from openpyxl import load_workbook

    plan = _plan()
    workbook = load_workbook(_built_tree() / "PCCM_stageA.xlsx")
    try:
        sheet = workbook[plan.sheet]
        expected = _cells(plan)
        for address, value in expected.items():
            assert sheet[address].value == value, f"{plan.sheet}!{address}"
        # AND NOTHING ELSE HOLDS A FORMULA. A stray formula anywhere on this
        # sheet is a second author.
        found = {cell.coordinate for row in sheet.iter_rows() for cell in row
                 if isinstance(cell.value, str) and cell.value.startswith("=")}
        planned = {address for address, value in expected.items()
                   if isinstance(value, str) and value.startswith("=")}
        assert found == planned, sorted(found ^ planned)[:10]
    finally:
        workbook.close()


def test_08_the_contracted_geometry_is_what_step_1_settled() -> None:
    plan, projection = _plan(), _projection()
    assert plan.row_window == 100
    assert projection["freeze_panes"] == "A6"
    assert plan.summary_row("overall_status") == 8
    assert [int(e["row"]) for e in plan.summary["rows"]] == list(range(8, 15))
    assert int(plan.register["header_row"]) == 17
    assert int(plan.register["first_row"]) == 18
    assert [c["column"] for c in plan.register["columns"]] == list("BCDEFG")
    assert [c["key"] for c in plan.register["columns"]] == [
        "check_id", "group", "severity", "subject", "message", "guidance"]
    assert plan.severity_order == ["ERROR", "WARNING", "INFO"]
    assert plan.group_order == ["Structure", "Inputs", "Calculation", "Simulation",
                                "Annual", "Sensitivity"]


def test_09_the_sheet_draws_no_chart_and_declares_no_button() -> None:
    from openpyxl import load_workbook

    workbook = load_workbook(_built_tree() / "PCCM_stageA.xlsx")
    try:
        sheet = workbook[_plan().sheet]
        assert not getattr(sheet, "_charts", []), "Model Check draws a chart"
        assert not getattr(sheet, "tables", {}), "Model Check declares a Table"
        assert not list(sheet.data_validations.dataValidation)
    finally:
        workbook.close()


def test_10_the_projection_is_emitted_and_refuses_a_missing_field() -> None:
    projection = _projection()
    validate_phase9_inspection(projection)
    for key in ("register", "advisory", "worksheet_safety", "vocabulary"):
        broken = copy.deepcopy(projection)
        del broken[key]
        with pytest.raises(ValueError, match=key):
            validate_phase9_inspection(broken)


def test_11_the_projection_carries_the_contract_threshold_not_its_own() -> None:
    projection, contract = _projection(), load_contract(CONTRACT_PATH)
    assert projection["advisory"]["threshold"] == contract.recommendation_for(
        ADVISORY_INPUT_KEY)
    assert projection["provenance"]["threshold_owner"].startswith("input_contract.yaml")


# ===========================================================================
# C. THE SCENARIO MATRIX
# ===========================================================================
def test_12_scenario_a_untouched_workbook() -> None:
    """WARNING from NOT CALCULATED. The live simulation INVALID is displayed
    faithfully as context and is NOT counted a second time; the optional
    publications are INFO."""
    result = _evaluate(_plan(), {"calculation_state": "NOT CALCULATED",
                                 "simulation_state": "INVALID"})
    assert _counts(result) == (0, 1)
    assert result["summary"]["overall_status"] == "WARNING"
    ids = _ids(result)
    assert "CAL-020" in ids, ids
    assert "SIM-020" in ids and "SIM-010" not in ids, ids
    assert "SIM-050" in ids and "ANN-030" in ids, ids
    severities = {row["check_id"]: row["severity"] for row in result["shown"]}
    assert severities["SIM-020"] == "INFO"
    assert severities["SIM-050"] == "INFO"
    assert severities["ANN-030"] == "INFO"


def test_13_scenario_a_the_advisory_only_when_below() -> None:
    plan = _plan()
    threshold = plan.recommended_iterations
    at = _evaluate(plan, {"simulation_state": "INVALID",
                          "requested_iterations": threshold})
    below = _evaluate(plan, {"simulation_state": "INVALID",
                             "requested_iterations": threshold - 1})
    assert "INP-010" not in _ids(at)
    assert "INP-010" in _ids(below)
    assert _counts(at) == (0, 1) and _counts(below) == (0, 2)


def test_14_scenario_b_calculated_and_never_simulated_is_pass() -> None:
    """A valid current model that has simply not been simulated is not
    defective, and Model Check does not report it as though it were."""
    result = _evaluate(_plan(), {"calculation_state": "CURRENT"})
    assert _counts(result) == (0, 0)
    assert result["summary"]["overall_status"] == "PASS"
    ids = _ids(result)
    for optional in ("SIM-050", "ANN-030", "SEN-010"):
        assert optional in ids, ids
    severities = {row["check_id"]: row["severity"] for row in result["shown"]}
    assert set(severities.values()) == {"INFO"}


def test_15_scenario_c_current_simulation_at_or_above_is_pass() -> None:
    plan = _plan()
    result = _evaluate(plan, {
        "calculation_state": "CURRENT", "simulation_state": "CURRENT",
        "requested_iterations": plan.recommended_iterations,
        "simulation_publication": "2026-01-01 00:00:00", "published_run_id": "RUN-1",
        "simulation_status_last_evaluated": "CURRENT"})
    assert _counts(result) == (0, 0)
    assert result["summary"]["overall_status"] == "PASS"
    assert "INP-010" not in _ids(result)


def test_16_scenario_d_current_simulation_below_is_one_warning() -> None:
    plan = _plan()
    result = _evaluate(plan, {
        "calculation_state": "CURRENT", "simulation_state": "CURRENT",
        "requested_iterations": plan.recommended_iterations - 1,
        "simulation_publication": "2026-01-01 00:00:00", "published_run_id": "RUN-1"})
    assert _counts(result) == (0, 1)
    warnings = [row for row in result["shown"] if row["severity"] == "WARNING"]
    assert [row["check_id"] for row in warnings] == ["INP-010"]
    assert warnings[0]["message"] == _projection()["advisory"]["message"]
    assert warnings[0]["guidance"] == _projection()["advisory"]["guidance"]


def test_17_scenario_e_request_drift() -> None:
    plan = _plan()
    result = _evaluate(plan, {
        "calculation_state": "CURRENT", "simulation_state": "STALE",
        "annual_distribution_state": "HISTORICAL", "annual_profile_state": "OTHER Px",
        "annual_profile_px": "P50",
        "simulation_publication": "2026-01-01 00:00:00", "published_run_id": "RUN-1",
        "simulation_status_last_evaluated": "CURRENT"})
    errors, warnings = _counts(result)
    assert errors == 0 and warnings >= 2, result["summary"]
    ids = _ids(result)
    for expected in ("SIM-030", "ANN-020", "ANN-050"):
        assert expected in ids, ids
    # THE PERSISTED ROW IS PRESENT AND LABELLED, and it is not counted.
    persisted = next(row for row in result["shown"] if row["check_id"] == "SIM-070")
    assert persisted["severity"] == "INFO"
    assert "last evaluated" in str(persisted["message"])


def test_18_scenario_f_an_invalid_model_does_not_double_count() -> None:
    """The calculation is the root cause and is counted once. The simulation
    INVALID that follows from it is context, not a second error."""
    result = _evaluate(_plan(), {"calculation_state": "INVALID",
                                 "simulation_state": "INVALID",
                                 "calculation_attempt_result": "REFUSED",
                                 "calculation_attempt_detail":
                                     "CL-0004 has mode above max."})
    errors, _ = _counts(result)
    assert errors == 1, result["summary"]
    assert result["summary"]["overall_status"] == "ERROR"
    ids = _ids(result)
    assert "CAL-010" in ids and "SIM-020" in ids and "SIM-010" not in ids
    # AND THE OWNER'S OWN NAMING OF THE OFFENDING INPUT IS SHOWN, as the
    # persisted attempt detail it is, labelled and uncounted.
    detail = next(row for row in result["shown"] if row["check_id"] == "CAL-051")
    assert detail["severity"] == "INFO"
    assert "CL-0004" in str(detail["subject"])


def test_19_scenario_f_structural_faults_are_the_root_cause_rows() -> None:
    result = _evaluate(_plan(), {"structural_report": _faults(3),
                                 "calculation_state": "INVALID"})
    errors, _ = _counts(result)
    assert errors == 4, result["summary"]
    structural = [row for row in result["shown"] if row["group"] == "Structure"]
    assert len(structural) == 3
    assert {row["check_id"] for row in structural} == {"CHK_ID_PATTERN"}
    assert all(str(row["subject"]).startswith("Cost Lines row") for row in structural)
    assert "STR-000" not in _ids(result)


def test_20_scenario_g_the_advisory_leaving_changes_no_other_row() -> None:
    plan = _plan()
    drift = {"calculation_state": "CURRENT", "simulation_state": "STALE",
             "simulation_publication": "2026-01-01 00:00:00", "published_run_id": "R"}
    below = _evaluate(plan, {**drift,
                             "requested_iterations": plan.recommended_iterations - 1})
    restored = _evaluate(plan, {**drift,
                                "requested_iterations": plan.recommended_iterations})
    assert _counts(below)[1] - _counts(restored)[1] == 1
    assert set(_ids(below)) - set(_ids(restored)) == {"INP-010"}
    assert set(_ids(restored)) - set(_ids(below)) == {"INP-011"}
    others = [i for i in _ids(below) if i not in ("INP-010", "INP-011")]
    assert others == [i for i in _ids(restored) if i not in ("INP-010", "INP-011")]


def test_21_scenario_h_many_errors_and_warnings_all_listed() -> None:
    plan = _plan()
    result = _evaluate(plan, {
        "structural_report": _faults(5), "calculation_state": "INVALID",
        "simulation_state": "STALE", "annual_profile_state": "HISTORICAL",
        "annual_distribution_state": "HISTORICAL",
        "requested_iterations": plan.recommended_iterations - 1,
        "simulation_publication": "x", "published_run_id": "R"})
    errors, warnings = _counts(result)
    assert result["summary"]["overall_status"] == "ERROR"
    shown = result["shown"]
    assert errors == len([r for r in shown if r["severity"] == "ERROR"])
    assert warnings == len([r for r in shown if r["severity"] == "WARNING"])
    assert int(result["summary"]["info_count"]) == len(
        [r for r in shown if r["severity"] == "INFO"])
    assert int(result["summary"]["total_checks"]) == len(shown)


def test_22_scenario_i_one_rule_twice_for_one_subject_is_one_row() -> None:
    """DEDUPLICATION BY check_id + subject, proved on the sheet's own formulas."""
    manifest = _manifest()
    checks = manifest["phase9_shell"]["model_check"]["checks"]
    original = next(c for c in checks if c["check_id"] == "CAL-020")
    twin = copy.deepcopy(original)
    twin["message"] = "A second declaration of the same issue for the same subject."
    checks.append(twin)
    plan = _plan(manifest)
    result = _evaluate(plan, {"calculation_state": "NOT CALCULATED"})
    assert _ids(result).count("CAL-020") == 1, _ids(result)
    assert _counts(result) == (0, 1)


def test_23_scenario_i_one_rule_on_two_subjects_is_two_rows() -> None:
    manifest = _manifest()
    checks = manifest["phase9_shell"]["model_check"]["checks"]
    original = next(c for c in checks if c["check_id"] == "CAL-020")
    twin = copy.deepcopy(original)
    twin["subject"] = "A different subject"
    checks.append(twin)
    plan = _plan(manifest)
    result = _evaluate(plan, {"calculation_state": "NOT CALCULATED"})
    assert _ids(result).count("CAL-020") == 2, _ids(result)
    assert _counts(result) == (0, 2)


def test_24_scenario_j_deterministic_order_regardless_of_declaration_order() -> None:
    """THE REGISTER IS SORTED AT BUILD TIME, so the order the manifest happens to
    list the checks in is not an input to anything."""
    plan = _plan()
    baseline = [c["check_id"] for c in plan.ordered_checks]
    manifest = _manifest()
    manifest["phase9_shell"]["model_check"]["checks"].reverse()
    shuffled = [c["check_id"] for c in _plan(manifest).ordered_checks]
    assert shuffled == baseline
    # AND THE ORDER IS THE CONTRACTED ONE.
    ranks = [(plan.severity_order.index(c["severity"]),
              plan.group_order.index(c["group"]), c["check_id"])
             for c in plan.ordered_checks]
    assert ranks == sorted(ranks)


def test_25_scenario_j_the_displayed_rows_follow_that_order() -> None:
    plan = _plan()
    result = _evaluate(plan, {
        "structural_report": _faults(2), "calculation_state": "INVALID",
        "simulation_state": "STALE",
        "requested_iterations": plan.recommended_iterations - 1})
    shown = result["shown"]
    ranks = [(plan.severity_order.index(str(row["severity"])),
              plan.group_order.index(str(row["group"]))) for row in shown]
    assert ranks == sorted(ranks), _ids(result)
    assert shown[0]["group"] == "Structure" and shown[0]["severity"] == "ERROR"


def test_26_scenario_k_more_than_a_hundred_issues() -> None:
    """THE WINDOW IS A RENDERING CAPACITY, NEVER A CLAIM ABOUT HOW MANY EXIST."""
    plan = _plan()
    total_faults = plan.row_window + 37
    result = _evaluate(plan, {"structural_report": _faults(total_faults),
                              "calculation_state": "INVALID"})
    summary = result["summary"]
    assert len(result["shown"]) == plan.row_window
    assert int(summary["error_count"]) >= total_faults
    declared_active = int(summary["total_checks"]) - total_faults
    assert declared_active > 0
    assert int(summary["total_checks"]) == total_faults + declared_active
    expected = str(plan.block["disclosure_template"]).format(
        window=plan.row_window, total=int(summary["total_checks"]))
    assert summary["disclosure"] == expected
    assert f"the first {plan.row_window} of" in expected
    # AND THE FIRST HUNDRED ARE THE DETERMINISTIC FIRST HUNDRED.
    assert all(row["group"] == "Structure" for row in result["shown"])


def test_27_scenario_k_the_disclosure_is_silent_when_it_fits() -> None:
    result = _evaluate(_plan(), {"calculation_state": "CURRENT"})
    assert result["summary"]["disclosure"] == ""
    assert int(result["summary"]["total_checks"]) <= _plan().row_window


def test_28_scenario_l_an_unused_slot_has_no_identity() -> None:
    """An empty slot is not a check. Not a blank, not a zero, not an empty
    string - #N/A, the accepted no-data representation."""
    result = _evaluate(_plan(), {"calculation_state": "CURRENT"})
    assert result["unused"], "every register row was used; the control proves nothing"
    for row in result["unused"]:
        for key, value in row.items():
            assert isinstance(value, ExcelError) and value == NA, (key, value)
    # A SHOWN ROW NEVER CARRIES A FABRICATED ZERO. That is the P8-3 hazard: an
    # empty string reaching a sheet through a value read as numeric 0, and then
    # being read as an identity. A blank SUBJECT, by contrast, is contracted -
    # Step 1 §5 says a model-wide issue has none - so it is permitted and only
    # the other five fields must carry text.
    for row in result["shown"]:
        for key, value in row.items():
            assert value != 0, (row["check_id"], key)
            if key != "subject":
                assert value != "", (row["check_id"], key)
    blank_subjects = [r for r in result["shown"] if r["subject"] == ""]
    assert blank_subjects, "no model-wide row was shown; the permission proves nothing"


def test_28b_a_truly_empty_candidate_cell_would_print_a_zero() -> None:
    """WHY THE BUILDER WRITES `=""` RATHER THAN LEAVING THE CELL ALONE. Excel
    reads an empty reference back as a hard zero, so a blank subject cell would
    reach the register as `0` - an identity nobody reported. Mutating the plan's
    empty-string subject back to a genuinely empty cell puts that zero on the
    sheet, which is what test_28 above refuses."""
    plan = _plan()
    cells = _cells(plan)
    column = plan.candidate_column("subject")
    emptied = 0
    for row in plan.declared_rows():
        if cells.get(f"{column}{row}") == '=""':
            cells[f"{column}{row}"] = None
            emptied += 1
    assert emptied, "no declared check has a model-wide subject; the mutation is inert"
    broken = _evaluate(plan, {"calculation_state": "CURRENT"}, cells=cells)
    assert any(row["subject"] == 0 for row in broken["shown"]), broken["shown"][:3]


def test_29_scenario_m_a_persisted_current_never_masquerades_as_live() -> None:
    """The live calculation state says STALE; the persisted rows still say what
    they were recorded as, and they are labelled and uncounted."""
    result = _evaluate(_plan(), {
        "calculation_state": "STALE", "simulation_state": "STALE",
        "calculation_attempt_result": "SUCCESS",
        "simulation_status_last_evaluated": "CURRENT",
        "simulation_publication": "x", "published_run_id": "R"})
    severities = {row["check_id"]: row["severity"] for row in result["shown"]}
    assert severities["CAL-030"] == "WARNING"
    assert "CAL-040" not in severities, "a live CURRENT row appeared for a stale model"
    assert severities["SIM-070"] == "INFO"
    persisted = next(r for r in result["shown"] if r["check_id"] == "SIM-070")
    assert "last evaluated" in str(persisted["message"])
    attempt = next(r for r in result["shown"] if r["check_id"] == "CAL-050")
    assert "last attempt" in str(attempt["message"])
    assert _counts(result)[0] == 0


def test_30_the_live_calculation_state_is_the_adapter_and_not_the_persisted_row() -> None:
    plan = _plan()
    entry = next(e for e in plan.readings["rows"] if e["key"] == "calculation_state")
    assert entry["procedure"] == "PCCM_ModelCheckCalculationState"
    formula = plan.reading_formula(entry)
    assert formula == "=PCCM_ModelCheckCalculationState()"
    assert "_Calc" not in formula and "PCCM_CalculationStatus" not in formula


# ===========================================================================
# D. THE ADVISORY BOUNDARY
# ===========================================================================
@pytest.mark.parametrize("offset,expected", [(-1, True), (0, False), (1, False)])
def test_31_the_advisory_boundary_is_strictly_less_than(offset: int,
                                                        expected: bool) -> None:
    plan = _plan()
    result = _evaluate(plan, {
        "calculation_state": "CURRENT",
        "requested_iterations": plan.recommended_iterations + offset})
    assert ("INP-010" in _ids(result)) is expected
    assert ("INP-011" in _ids(result)) is not expected


def test_32_the_advisory_is_a_warning_and_never_a_refusal() -> None:
    advisory = _projection()["advisory"]
    assert advisory["severity"] == "WARNING"
    assert advisory["refuses"] is False
    assert advisory["comparison"] == "strictly less than"
    # AND IT COEXISTS WITH AN ERROR RATHER THAN BEING SUPPRESSED BY ONE.
    plan = _plan()
    result = _evaluate(plan, {"calculation_state": "INVALID",
                              "requested_iterations": plan.recommended_iterations - 1})
    assert "INP-010" in _ids(result)
    assert result["summary"]["overall_status"] == "ERROR"


def test_33_the_advisory_is_live_and_not_publication_based() -> None:
    """It is advice about the REQUEST. A stale simulation does not change what
    the operator is asking for next time, and no publication is needed for the
    advice to be true."""
    plan = _plan()
    below = plan.recommended_iterations - 1
    without = _evaluate(plan, {"calculation_state": "CURRENT",
                               "requested_iterations": below})
    stale = _evaluate(plan, {"calculation_state": "CURRENT",
                             "simulation_state": "STALE",
                             "published_iterations_run": plan.recommended_iterations * 10,
                             "simulation_publication": "x", "published_run_id": "R",
                             "requested_iterations": below})
    assert "INP-010" in _ids(without) and "INP-010" in _ids(stale)
    entry = next(e for e in plan.readings["rows"] if e["key"] == "requested_iterations")
    assert entry["kind"] == "defined_name"
    assert entry["defined_name"] == "inpMonteCarloIterations"


def test_34_the_advisory_wording_is_the_contracted_wording() -> None:
    plan, projection = _plan(), _projection()
    formatted = f"{plan.recommended_iterations:,}"
    assert projection["advisory"]["message"] == (
        f"Monte Carlo iterations are below the recommended {formatted}; "
        "simulation precision may be lower at this setting.")
    assert projection["advisory"]["guidance"] == (
        f"Increase Monte Carlo Iterations on Setup to {formatted} or more before "
        "the next simulation run.")


# ===========================================================================
# E. PRECEDENCE AND THE ACTIONABLE COUNT
# ===========================================================================
def test_35_info_never_moves_the_overall_status() -> None:
    plan = _plan()
    formula = plan.summary_formula("overall_status")
    assert plan.summary_cell("info_count") not in formula, formula
    assert plan.summary_cell("error_count") in formula
    assert plan.summary_cell("warning_count") in formula
    result = _evaluate(plan, {"calculation_state": "CURRENT"})
    assert int(result["summary"]["info_count"]) > 0
    assert result["summary"]["overall_status"] == "PASS"


def test_36_the_precedence_is_error_then_warning_then_pass() -> None:
    plan = _plan()
    cases = [
        ({"calculation_state": "INVALID"}, "ERROR"),
        ({"calculation_state": "STALE"}, "WARNING"),
        ({"calculation_state": "CURRENT"}, "PASS"),
    ]
    for readings, expected in cases:
        result = _evaluate(plan, readings)
        assert result["summary"]["overall_status"] == expected, readings
    # AND AN ERROR OUTRANKS A WARNING RATHER THAN HIDING IT.
    both = _evaluate(plan, {"calculation_state": "INVALID", "simulation_state": "STALE",
                            "requested_iterations": plan.recommended_iterations - 1})
    assert both["summary"]["overall_status"] == "ERROR"
    assert int(both["summary"]["warning_count"]) > 0


def test_37_the_counts_reconcile_exactly_to_the_register() -> None:
    plan = _plan()
    for readings in (
        {"calculation_state": "NOT CALCULATED", "simulation_state": "INVALID"},
        {"calculation_state": "CURRENT", "simulation_state": "CURRENT",
         "simulation_publication": "x", "published_run_id": "R"},
        {"structural_report": _faults(7), "calculation_state": "INVALID"},
        {"calculation_state": "STALE", "annual_profile_state": "HISTORICAL",
         "requested_iterations": plan.recommended_iterations - 1},
    ):
        result = _evaluate(plan, readings)
        shown = result["shown"]
        summary = result["summary"]
        for word, key in (("ERROR", "error_count"), ("WARNING", "warning_count"),
                          ("INFO", "info_count")):
            assert int(summary[key]) == len(
                [r for r in shown if r["severity"] == word]), (readings, word)
        assert int(summary["total_checks"]) == len(shown), readings


def test_38_the_optional_publications_are_never_warnings() -> None:
    plan = _plan()
    severities = {c["check_id"]: c["severity"] for c in plan.ordered_checks}
    for optional in ("SIM-050", "ANN-030"):
        assert severities[optional] == "INFO", optional
    result = _evaluate(plan, {"calculation_state": "CURRENT"})
    assert _counts(result) == (0, 0)


def test_38b_the_projection_names_the_optional_publications() -> None:
    """§11 REQUIRES A RUNNER TO BE ABLE TO FIND THEM. The register lists them by
    id so a Windows runner never has to know which ids they are, and the
    validator refuses a projection in which one of them is actionable."""
    projection = _projection()
    optional = projection["register"]["optional_publications"]
    assert set(optional) == {"SIM-050", "ANN-030", "SEN-010"}, optional
    severities = {c["check_id"]: c["severity"] for c in _plan().ordered_checks}
    for check_id in optional:
        assert severities[check_id] == _plan().informational, check_id
    # AND THE VALIDATOR REFUSES ONE THAT WAS PROMOTED.
    broken = copy.deepcopy(projection)
    for entry in broken["evaluation"]["declared_checks"]:
        if entry["check_id"] == "SIM-050":
            entry["severity"] = "WARNING"
    with pytest.raises(ValueError, match="not a defect"):
        validate_phase9_inspection(broken)


def test_38c_the_loader_refuses_an_actionable_optional_publication() -> None:
    from pccm_builder.spec_loader import SpecError

    def edit(block):
        for check in block["checks"]:
            if check.get("optional_publication"):
                check["severity"] = "WARNING"
    with pytest.raises(SpecError, match="not a defect"):
        _mutate_manifest(edit)


def test_39_not_calculated_is_actionable() -> None:
    plan = _plan()
    severities = {c["check_id"]: c["severity"] for c in plan.ordered_checks}
    assert severities["CAL-020"] == "WARNING"
    assert "WARNING" in plan.actionable


def test_40_the_two_state_axes_stay_apart() -> None:
    """P8-3 cost a Windows run to the assumption that they were one axis. They
    are two rows, two vocabularies and two readings."""
    plan = _plan()
    calc_rows = [c for c in plan.ordered_checks if c["group"] == "Calculation"]
    sim_rows = [c for c in plan.ordered_checks if c["group"] == "Simulation"]
    assert calc_rows and sim_rows
    calc_reading = plan.reading_cell("calculation_state")
    sim_reading = plan.reading_cell("simulation_state")
    assert calc_reading != sim_reading
    # A BLANK SIMULATION IS NOT INVALID, AND INVALID IS NEVER RENDERED BLANK.
    blank = _evaluate(plan, {"calculation_state": "CURRENT", "simulation_state": ""})
    invalid = _evaluate(plan, {"calculation_state": "CURRENT",
                               "simulation_state": "INVALID"})
    assert "SIM-010" not in _ids(blank)
    assert "SIM-010" in _ids(invalid)
    assert _counts(blank) == (0, 0) and _counts(invalid) == (1, 0)


# ===========================================================================
# E2. P9-2A - THE LIVE REFUSAL DETAIL
# ===========================================================================
# WHAT CHANGED AND WHY. The actionable calculation ERROR used to carry a static
# contract sentence, and the only text naming the fault was the PERSISTED last
# attempt - history, blank until somebody has pressed Calculate, and stale the
# moment the model moves past it. The row now carries the sentence the CURRENT
# preparation wrote on its way to deciding the status.
#
# WHAT DID NOT CHANGE, AND IS A DECLARED GAP. Step-1 §11 scenario F asks that
# row to name the offending PERMANENT ID in its Subject. The owner has the id in
# hand when it refuses, but publishes only a sentence: there is no structured
# field anywhere in the accepted owners carrying the offending id on its own.
# Recovering it would mean parsing prose, which is refused. test_46c states that
# limitation as a fact about the source so it cannot be forgotten.

def _refusal(driver: str = "CL-0001") -> str:
    """A refusal sentence shaped exactly the way the owner shapes one.

    modCalcCheck.DriverLabel builds "cost line <id>" or "risk <id>", and
    OrderingFailure appends ": <Distribution> requires Min <= Most Likely <= Max".
    """
    return f"cost line {driver}: Triangular requires Min <= Most Likely <= Max"


def test_46a_the_actionable_error_carries_the_live_reason() -> None:
    """A. THE LIVE SENTENCE REACHES THE ACTIONABLE ROW, and the persisted one is
    somewhere else entirely."""
    plan = _plan()
    result = _evaluate(plan, {
        "calculation_state": "INVALID", "simulation_state": "INVALID",
        "calculation_refusal_detail": _refusal("CL-0001"),
        "calculation_attempt_result": "REFUSED",
        "calculation_attempt_detail": _refusal("CL-0009")})
    error = next(r for r in result["shown"] if r["severity"] == "ERROR")
    assert error["check_id"] == "CAL-010", error
    assert error["group"] == "Calculation"
    assert error["message"] == _refusal("CL-0001"), error["message"]
    # AND IT IS NOT THE PERSISTED ONE. The two are deliberately different here,
    # so a row that reached for history instead of the live model would show it.
    assert _refusal("CL-0009") not in str(error["message"])
    persisted = next(r for r in result["shown"] if r["check_id"] == "CAL-051")
    assert persisted["subject"] == _refusal("CL-0009")
    assert persisted["severity"] == plan.informational
    assert "last attempt" in str(persisted["message"])


def test_46b_the_live_reason_moves_with_the_model_and_clears_with_it() -> None:
    """B and C. The live row follows the CURRENT model - no publication, no
    calculation attempt, nothing persisted is involved - and it is gone the
    moment the model is valid again."""
    plan = _plan()
    first = _evaluate(plan, {"calculation_state": "INVALID",
                             "calculation_refusal_detail": _refusal("CL-0001")})
    moved = _evaluate(plan, {"calculation_state": "INVALID",
                             "calculation_refusal_detail": _refusal("CL-0002")})
    assert next(r for r in first["shown"] if r["check_id"] == "CAL-010")["message"] == \
        _refusal("CL-0001")
    assert next(r for r in moved["shown"] if r["check_id"] == "CAL-010")["message"] == \
        _refusal("CL-0002")
    # NOTHING WAS PUBLISHED AND NOTHING WAS ATTEMPTED in either reading, so the
    # live row cannot have come from a persisted source.
    for result in (first, moved):
        assert "CAL-050" not in _ids(result) and "CAL-051" not in _ids(result)
    # AND CORRECTING THE INPUT CLEARS IT.
    corrected = _evaluate(plan, {"calculation_state": "CURRENT",
                                 "calculation_refusal_detail": ""})
    assert "CAL-010" not in _ids(corrected), _ids(corrected)
    assert _counts(corrected) == (0, 0)


def test_46c_no_prose_is_parsed_anywhere() -> None:
    """THE RULE THE WHOLE CORRECTION EXISTS TO KEEP. The permanent id is now a
    VALUE, threaded out of the owners that already hold it. Nothing recovers it
    from the sentence beside it - not in VBA, not in a formula, not here.

    Three modules write refusal sentences in four different shapes, so a parser
    would have been broken by a fifth. There is none.
    """
    plan = _plan()
    # NOT IN THE WORKSHEET. The only text-splitting on the surface is the
    # structural report's own line separator, which its owner writes.
    for reading in ("calculation_refusal_detail", "calculation_attempt_detail"):
        cell = plan.reading_cell(reading)
        for value in _cells(plan).values():
            if isinstance(value, str) and cell in value:
                for splitter in ("FIND(", "MID(", "LEFT(", "RIGHT(", "SEARCH(", "SUBSTITUTE("):
                    assert splitter not in value, (
                        f"a Model Check formula takes a refusal sentence apart: {value}")
    # NOR IN THE ADAPTERS. They hand a ByRef out; they do not read one.
    for name in ("PCCM_ModelCheckRefusalSubject", "PCCM_ModelCheckRefusalDetail"):
        body = _procedure("modResultsState", name)
        for splitter in ("InStr", "Mid$", "Mid(", "Left(", "Right(", "Split("):
            assert splitter not in body, f"{name} parses text: {splitter}"
    # NOR IN THE OWNERS. Every subject assignment is a permanent id read from a
    # resolved driver, never a fragment of a sentence.
    import re

    for module in ("modCalcCheck", "modCalcResolve", "modCalcReport",
                   "modCalcAnalytical"):
        try:
            _assert_sources_are_permanent_ids(_code(module))
        except AssertionError as failure:
            raise AssertionError(f"{module}: {failure}") from None


def test_46c1_every_subject_names_the_same_driver_the_message_does() -> None:
    """THE WRONG-DRIVER MUTATION, REFUSED STRUCTURALLY. A subject taken from a
    different index than the sentence beside it would name an innocent driver -
    worse than naming none - and would read perfectly plausibly. So the id
    EXPRESSION must be one the owner already uses for that same refusal."""
    import re

    expressions: dict[str, set[str]] = {}
    for module in ("modCalcCheck", "modCalcResolve", "modCalcAnalytical",
                   "modCalcReport"):
        code = _code(module)
        used = set(re.findall(r"([A-Za-z_][\w.()+ ]*?\.PermanentId)", code))
        assigned = {source for source in _subject_sources(code)
                    if source.endswith(".PermanentId")}
        assert assigned, f"{module} threads no subject at all"
        unknown = {a for a in assigned if a not in used}
        assert not unknown, (
            f"{module} sets subject from an id expression it uses nowhere else: {unknown}")
        expressions[module] = assigned
    # AND THE ONE THAT MATTERS IS THE LOOP'S OWN INDEX, not a neighbour's.
    check = _code("modCalcCheck")
    loop = check[check.index("For index = 0 To model.DriverCount - 1"):]
    loop = loop[:loop.index("Next index")]
    assert "subject = model.Drivers(LBound(model.Drivers) + index).PermanentId" in loop
    assert "index + 1" not in loop and "index - 1" not in loop, loop


def test_46c2_the_owners_clear_the_subject_so_it_cannot_leak() -> None:
    """THE LEAK THAT WOULD BE INVISIBLE. A subject set for driver seven and left
    behind would attach the next model-wide refusal - a missing register, an
    unusable discount rate - to a driver that had nothing to do with it.

    Two rules make it impossible: the preparation clears at entry, and every
    owner that sets a subject speculatively clears it again on its own success.
    """
    prepare = _procedure("modCalcReport", "PrepareCurrentCalculation")
    assert "subject = vbNullString" in prepare, "the preparation does not clear at entry"
    assert prepare.index("subject = vbNullString") < prepare.index("ResolveModel"), (
        "the clear happens after the first owner could have set it")

    # AND EVERY TRAVERSAL ENTRY CLEARS BEFORE IT LOOKS AT ANYTHING. The three
    # procedures a whole traversal starts at take the parameter from a caller
    # and must not inherit whatever that caller was holding. The clears further
    # down are what stop a SUCCESSFUL owner leaving an id behind; these are what
    # stop one traversal's id reaching the next, and they are separate rules -
    # dropping either used to be visible only as a moved byte pin.
    for module, name, first in (("modCalcReport", "PrepareCurrentCalculation", "ResolveModel"),
                                ("modCalcResolve", "ResolveModel", "ResolveDrivers"),
                                ("modCalcCheck", "CheckResolvedModel", "For ")):
        body = _procedure(module, name)
        assert "subject = vbNullString" in body, (
            f"{module}.{name} takes a subject from its caller and never clears it")
        assert body.index("subject = vbNullString") < body.index(first), (
            f"{module}.{name} clears the subject only after it has begun work, so a "
            "caller's id can still reach a refusal raised before that point")
    for module, name in (("modCalcCheck", "CheckResolvedModel"),
                         ("modCalcResolve", "ResolveModel"),
                         ("modCalcResolve", "ReadDriverRow"),
                         ("modCalcResolve", "ResolveProfileWeights")):
        body = _procedure(module, name)
        if "subject = vbNullString" not in body and ".PermanentId" not in body:
            continue
        lines = [line.strip() for line in body.splitlines()]
        sets = [i for i, line in enumerate(lines) if line.startswith("subject = ")
                and not line.endswith("vbNullString")]
        clears = [i for i, line in enumerate(lines) if line == "subject = vbNullString"]
        success = [i for i, line in enumerate(lines) if line == f"{name} = True"]
        if not sets:
            continue
        assert clears, f"{module}.{name} sets a subject and never clears one"
        assert success, f"{module}.{name} has no single success line to guard"
        # A CLEAR STANDS BETWEEN THE LAST SPECULATIVE SET AND THE SUCCESS EXIT.
        assert any(max(sets) < clear < success[-1] for clear in clears), (
            f"{module}.{name} can succeed while still holding a subject")


def test_46c3_no_owner_keeps_a_last_refusal_cache() -> None:
    """NO STATIC, NO MODULE STATE. The subject is an output of one call, not a
    thing anybody remembers between calls."""
    import re

    for module in ("modCalcCheck", "modCalcResolve", "modCalcReport",
                   "modCalcAnalytical", "modResultsState"):
        code = _code(module)
        assert not re.search(r"^\s*Static\b", code, re.M), f"{module} declares a Static"
        # MODULE LEVEL IS COLUMN ZERO AND ABOVE THE FIRST PROCEDURE. A `Dim`
        # inside a procedure body is a local and is none of this control's
        # business; scanning for one indented would convict every function in
        # the file, which is how the first draft of this failed.
        for line in code.splitlines():
            if re.match(r"^(Public|Private)\s+(Function|Sub)\b", line):
                break
            assert not re.match(r"^(Public|Private|Dim)\s+\w+\s+As\s", line), (
                f"{module} declares module-level mutable state: {line.strip()}")


def test_46d_the_root_cause_is_still_counted_exactly_once() -> None:
    """The live reason did not add a row. One actionable ERROR, the simulation
    that follows from it still context, and the counts still reconcile."""
    plan = _plan()
    result = _evaluate(plan, {
        "calculation_state": "INVALID", "simulation_state": "INVALID",
        "calculation_refusal_detail": _refusal(),
        "calculation_attempt_result": "REFUSED",
        "calculation_attempt_detail": _refusal("CL-0009")})
    errors = [r for r in result["shown"] if r["severity"] == "ERROR"]
    assert len(errors) == 1, [r["check_id"] for r in errors]
    assert _counts(result) == (1, 0)
    assert result["summary"]["overall_status"] == "ERROR"
    assert int(result["summary"]["total_checks"]) == len(result["shown"])
    assert "SIM-020" in _ids(result) and "SIM-010" not in _ids(result)


def test_46e_mutation_the_persisted_detail_substituted_for_the_live_one() -> None:
    """E. THE SUBSTITUTION THAT WOULD LOOK RIGHT AND BE WRONG. History reads
    plausibly; it is simply about a workbook that no longer exists."""
    def edit(block):
        for check in block["checks"]:
            if check["check_id"] == "CAL-010":
                check["message"] = "{calculation_attempt_detail}"
    plan = _mutate_manifest(edit)
    broken = _evaluate(plan, {"calculation_state": "INVALID",
                              "calculation_refusal_detail": _refusal("CL-0001"),
                              "calculation_attempt_detail": _refusal("CL-0009")})
    message = next(r for r in broken["shown"] if r["check_id"] == "CAL-010")["message"]
    assert message == _refusal("CL-0009"), "the mutation changed nothing"
    with pytest.raises(AssertionError):
        _assert_live_reason(plan)


def _assert_live_reason(plan: ModelCheckPlan) -> None:
    check = next(c for c in plan.ordered_checks if c["check_id"] == "CAL-010")
    assert str(check["message"]) == "{calculation_refusal_detail}", check["message"]


def test_46f_mutation_the_live_reason_is_blanked() -> None:
    def edit(block):
        for check in block["checks"]:
            if check["check_id"] == "CAL-010":
                check["message"] = "The model is invalid."
    plan = _mutate_manifest(edit)
    broken = _evaluate(plan, {"calculation_state": "INVALID",
                              "calculation_refusal_detail": _refusal()})
    assert next(r for r in broken["shown"]
                if r["check_id"] == "CAL-010")["message"] == "The model is invalid."
    with pytest.raises(AssertionError):
        _assert_live_reason(plan)
    _assert_live_reason(_plan())


def test_46g_mutation_a_duplicate_actionable_error_for_the_same_root_cause() -> None:
    def edit(block):
        for check in block["checks"]:
            if check["check_id"] == "SIM-020":
                check["severity"] = "ERROR"
    plan = _mutate_manifest(edit)
    broken = _evaluate(plan, {"calculation_state": "INVALID", "simulation_state": "INVALID",
                              "calculation_refusal_detail": _refusal()})
    assert _counts(broken) == (2, 0), "the mutation changed nothing"
    healthy = _evaluate(_plan(), {"calculation_state": "INVALID",
                                  "simulation_state": "INVALID",
                                  "calculation_refusal_detail": _refusal()})
    assert _counts(healthy) == (1, 0)


def test_46h_mutation_a_persisting_path_supplies_the_live_reason() -> None:
    """D and E. The live reason may not come from a function that writes."""
    def edit(block):
        for entry in block["evaluation"]["readings"]["rows"]:
            if entry["key"] == "calculation_refusal_detail":
                entry["procedure"] = "PCCM_CalculationStatus"
    plan = _mutate_manifest(edit)
    inspection = build_phase9_inspection(plan)
    with pytest.raises(ValueError, match="persists"):
        validate_phase9_inspection(inspection)


def test_46i_the_live_reason_path_is_read_only_and_volatile() -> None:
    """D. The transitive walk in test_41 covers every cell-called procedure; this
    names what the new one is, so a reader can see the split without re-deriving
    it."""
    plan = _plan()
    entry = next(e for e in plan.readings["rows"]
                 if e["key"] == "calculation_refusal_detail")
    assert entry["procedure"] == "PCCM_ModelCheckRefusalDetail"
    assert plan.reading_formula(entry) == "=PCCM_ModelCheckRefusalDetail()"
    adapter = _procedure("modResultsState", "PCCM_ModelCheckRefusalDetail")
    assert "Application.Volatile True" in adapter
    assert "modCalcReport.CalcReportDerivedStatus(detail, subject)" in adapter
    assert "CVErr(xlErrValue)" in adapter, "the adapter must fail loud, not wrong"
    assert "PCCM_CalculationStatus" not in adapter
    assert "PCCM_CalculationAttemptDetail" not in adapter, (
        "the live reason may not be taken from the persisted attempt")
    # AND THE OWNER HANDS BACK WHAT IT ALREADY WROTE; nothing is re-derived.
    exposure = _procedure("modCalcReport", "CalcReportDerivedStatus")
    assert "ByRef detail As String, ByRef subject As String" in exposure
    assert "Optional" not in exposure, (
        "a typed Optional with no default is a VBA compile error, and every "
        "caller supplies the argument")
    assert "PrepareCurrentCalculation(package, detail, subject)" in exposure
    assert "WriteStatusBlock" not in exposure


# ===========================================================================
# E3. P9-2B - THE STRUCTURED REFUSAL SUBJECT
# ===========================================================================
# THE REFUSAL FAMILIES, AND WHICH OF THEM CAN NAME A DRIVER.
#
# This is the audit §3 asks for, written as data so it can be checked rather
# than believed. A family is driver-specific when the refusing owner holds a
# ResolvedDriver at the moment it refuses; it is model-wide when no driver is at
# fault and no id exists to report.
REFUSAL_FAMILIES = (
    # (owner, family, driver-specific, where the subject comes from)
    ("modCalcResolve", "structural prerequisites", False, ""),
    ("modCalcResolve", "applied timeline / project years", False, ""),
    ("modCalcResolve", "register missing", False, ""),
    ("modCalcResolve", "driver row: identity unreadable", False,
     "the Permanent ID itself could not be read, so there is no id to name"),
    ("modCalcResolve", "driver row: currency, profile, distribution, scalars",
     True, "ReadDriverRow, from driver.PermanentId once the id is known"),
    ("modCalcResolve", "FX table / reporting currency", False, ""),
    ("modCalcResolve", "inflation grid / profile years", False, ""),
    ("modCalcResolve", "profiling grid row and columns", True,
     "ResolveProfileWeights, from drivers(...).PermanentId at the loop"),
    ("modCalcResolve", "driver currency not in reference set", True,
     "AttachDriverFx, from model.Drivers(index).PermanentId"),
    ("modCalcCheck", "timeline / discount rate / driver count", False, ""),
    ("modCalcCheck", "three-point ordering (cost line and risk)", True,
     "CheckResolvedModel, from model.Drivers(...).PermanentId at the loop"),
    ("modCalcCheck", "Quantity, Probability", True,
     "CheckResolvedModel, same loop"),
    ("modCalcCheck", "profiling weight sum", True,
     "CheckResolvedModel, same loop"),
    # CLOSED AT P9-3. These four were reported at P9-2B as named non-coverage,
    # because both modules were at their raw-line ceiling and no line could be
    # spared. They are threaded now, and the ceilings did not move: every
    # assignment rides on a statement that was already there.
    ("modCalcAnalytical", "per-driver conditioning magnitude (Contribute)",
     True, "AccumulateTotals, from audits(index).PermanentId at both passes"),
    ("modCalcAnalytical",
     "per-driver-per-year annual conditioning magnitude (RecordAnnual)", True,
     "BuildAnnualSeries, from drivers(...).PermanentId at the per-year pass"),
    ("modCalcAnalytical", "I5 profile-sum identity", True,
     "Reconcile, from drivers(index).PermanentId at the I5 loop"),
    ("modCalcAnalytical", "measure totals, conditioning coefficient", False, ""),
    ("modCalcAnalytical", "annual series, applied timeline", False, ""),
    ("modCalcReport", "inflation profile not in reference set", True,
     "BuildDriverFactors, from the id the loop has just copied"),
    ("modCalcReport", "Knom / Kpv factor build", True,
     "BuildDriverFactors, same loop"),
    ("modCalcReport", "driver audit build", True,
     "BuildAudits, from package.Drivers(index).PermanentId at the loop"),
    ("modCalcReport", "fingerprint record encoding", True,
     "BuildFingerprint, from package.Model.Drivers(index).PermanentId at the loop"),
    ("modCalcReport", "reconciliation identities, fingerprint construction", False, ""),
    ("modCalcReport", "factor tables, discount factors", False, ""),
)


_THREADED = {"modCalcResolve", "modCalcCheck", "modCalcAnalytical",
             "modCalcReport"}


def _subject_statements(text: str) -> list[str]:
    """Every `subject = ...` statement in *text*, wherever it sits on its line.

    P9-3 threads two modules that had no raw line to spare, so their assignments
    ride on statements that were already there - `Next slot: subject =
    vbNullString`. A scan anchored to the start of a line would see none of them
    and would pass by seeing nothing, which is the worst way for a control to
    be green.
    """
    found: list[str] = []
    for line in text.splitlines():
        for part in line.split(":"):
            part = part.strip()
            if part.startswith("subject = "):
                found.append(part)
    return found


def test_46j_every_threaded_owner_really_threads_and_none_is_left_out() -> None:
    """THE COVERAGE TABLE IS CHECKED AGAINST THE SOURCE, not merely written.

    A module the table says threads must actually set a subject; a module it
    says does not must actually not - so the table cannot quietly become a
    description of what somebody hoped was true.
    """
    for module in ("modCalcResolve", "modCalcCheck", "modCalcAnalytical",
                   "modCalcReport"):
        sets = _subject_statements(_code(module))
        if module in _THREADED:
            assert sets, f"{module} is listed as threaded and threads nothing"
        else:
            assert not sets, (
                f"{module} threads a subject; the coverage table says it does not")

    # AND NOTHING IS STILL REPORTED AS UNCOVERED. P9-2B named two owners it
    # could not reach; P9-3 reached them, so the table may no longer carry a
    # single driver-specific family without a structured source. If a later
    # phase adds one, this fails rather than letting the gap reappear quietly.
    unsourced = [(owner, family) for owner, family, specific, source
                 in REFUSAL_FAMILIES
                 if specific and (source.startswith("NOT THREADED") or not source)]
    assert not unsourced, (
        f"a driver-specific refusal family has no structured subject: {unsourced}")

    # AND EVERY OWNER THE TABLE NAMES IS ONE THAT EXISTS AND IS THREADED.
    assert {owner for owner, _f, specific, _s in REFUSAL_FAMILIES
            if specific} == _THREADED | {"modCalcResolve"}, sorted(
        {owner for owner, _f, specific, _s in REFUSAL_FAMILIES if specific})


@pytest.mark.parametrize("module,procedure,expression", [
    ("modCalcAnalytical", "AccumulateTotals", "who"),
    ("modCalcAnalytical", "BuildAnnualSeries",
     "drivers(LBound(drivers) + order(slot)).PermanentId"),
    ("modCalcAnalytical", "Reconcile", "drivers(index).PermanentId"),
    ("modCalcReport", "BuildDriverFactors",
     "package.Model.Drivers(index).PermanentId"),
    ("modCalcReport", "BuildAudits", "package.Drivers(index).PermanentId"),
    ("modCalcReport", "BuildFingerprint",
     "package.Model.Drivers(index).PermanentId"),
])
def test_46v_each_remaining_owner_sets_the_id_its_own_loop_already_held(
        module: str, procedure: str, expression: str) -> None:
    """P9-3, ONE FAMILY AT A TIME. Each of the six procedures that reach a
    driver-specific refusal assigns the permanent id the enclosing loop had in
    hand - not a lookup, not a second traversal, not a reconstruction."""
    body = _procedure(module, procedure)
    assert f"subject = {expression}" in body, (module, procedure, expression)
    # AND THE ID IS SET BEFORE ANY REFUSAL IN THAT PROCEDURE CAN FIRE, so a
    # refusal never reports the previous driver.
    statements = [part.strip() for line in body.splitlines()
                  for part in line.split(":")]
    first_set = next(i for i, s in enumerate(statements)
                     if s.startswith("subject = ") and s != "subject = vbNullString")
    # A REFUSAL IS EITHER A SENTENCE WRITTEN HERE OR A DELEGATE'S FALSE. In
    # AccumulateTotals and BuildFingerprint the per-driver refusal is Contribute
    # or DriverRecord returning False and this procedure leaving; the sentence is
    # the callee's. Both shapes count, or the control would only see one of them.
    refusals = [i for i, s in enumerate(statements)
                if (s.startswith("detail = ") and s != "detail = vbNullString")
                or "Exit Function" in s]
    driver_refusals = [i for i in refusals if i > first_set]
    assert driver_refusals, (
        f"{module}.{procedure} sets a subject no refusal below it can use")


@pytest.mark.parametrize("module,procedure", [
    ("modCalcAnalytical", "AccumulateTotals"),
    ("modCalcAnalytical", "BuildAnnualSeries"),
    ("modCalcAnalytical", "Reconcile"),
    ("modCalcReport", "BuildDriverFactors"),
    ("modCalcReport", "BuildAudits"),
    ("modCalcReport", "BuildFingerprint"),
])
def test_46w_each_remaining_owner_clears_before_its_model_wide_phase(
        module: str, procedure: str) -> None:
    """A measure total, a fingerprint construction, an annual series: none of
    them is about one driver, and none may inherit the last id the per-driver
    loop above them happened to leave behind."""
    body = _procedure(module, procedure)
    statements = [part.strip() for line in body.splitlines()
                  for part in line.split(":")]
    sets = [i for i, s in enumerate(statements)
            if s.startswith("subject = ") and s != "subject = vbNullString"]
    clears = [i for i, s in enumerate(statements) if s == "subject = vbNullString"]
    assert sets, (module, procedure)
    assert any(clear > max(sets) for clear in clears), (
        f"{module}.{procedure} can leave a driver id set once its per-driver "
        "phase is over")


@pytest.mark.parametrize("driver", ["CL-0001", "RSK-0004"])
def test_46k_the_actionable_error_names_the_offending_driver(driver: str) -> None:
    """A, B, C, D. Whatever family refused, the row shows the id as a VALUE in
    the Subject column - a cost line or a risk, from any threaded owner."""
    plan = _plan()
    result = _evaluate(plan, {
        "calculation_state": "INVALID", "simulation_state": "INVALID",
        "calculation_refusal_detail": f"cost line {driver}: Triangular requires "
                                      "Min <= Most Likely <= Max",
        "calculation_refusal_subject": driver})
    error = next(r for r in result["shown"] if r["severity"] == "ERROR")
    assert error["check_id"] == "CAL-010"
    assert error["subject"] == driver, error
    assert driver in str(error["message"])
    assert _counts(result) == (1, 0)


def test_46l_a_model_wide_refusal_leaves_the_subject_blank() -> None:
    """F. No driver is at fault in a missing register or an unusable discount
    rate, and a fabricated id would be worse than none."""
    result = _evaluate(_plan(), {
        "calculation_state": "INVALID",
        "calculation_refusal_detail": "the FX table tblFXRates is missing",
        "calculation_refusal_subject": ""})
    error = next(r for r in result["shown"] if r["severity"] == "ERROR")
    assert error["subject"] == "", error
    assert error["subject"] != 0, "a blank subject reached the sheet as a zero"
    assert "tblFXRates" in str(error["message"])
    assert _counts(result) == (1, 0)


def test_46m_the_subject_moves_and_clears_with_the_model() -> None:
    """G, H and I. It follows the CURRENT model; nothing published or attempted
    is consulted, and a persisted id that says otherwise does not win."""
    plan = _plan()
    first = _evaluate(plan, {"calculation_state": "INVALID",
                             "calculation_refusal_detail": _refusal("CL-0001"),
                             "calculation_refusal_subject": "CL-0001"})
    moved = _evaluate(plan, {"calculation_state": "INVALID",
                             "calculation_refusal_detail": _refusal("CL-0002"),
                             "calculation_refusal_subject": "CL-0002",
                             # HISTORY SAYS SOMETHING ELSE ENTIRELY, on purpose.
                             "calculation_attempt_result": "REFUSED",
                             "calculation_attempt_detail": _refusal("CL-0009")})
    assert next(r for r in first["shown"] if r["check_id"] == "CAL-010")["subject"] == "CL-0001"
    row = next(r for r in moved["shown"] if r["check_id"] == "CAL-010")
    assert row["subject"] == "CL-0002", row
    persisted = next(r for r in moved["shown"] if r["check_id"] == "CAL-051")
    assert "CL-0009" in str(persisted["subject"])
    assert persisted["severity"] == plan.informational
    corrected = _evaluate(plan, {"calculation_state": "CURRENT",
                                 "calculation_refusal_detail": "",
                                 "calculation_refusal_subject": ""})
    assert "CAL-010" not in _ids(corrected)
    assert _counts(corrected) == (0, 0)


def test_46n_the_subject_reading_is_the_live_adapter() -> None:
    """§5 wiring, and §6's read-only rule for the path it adds."""
    plan = _plan()
    entry = next(e for e in plan.readings["rows"]
                 if e["key"] == "calculation_refusal_subject")
    assert entry["procedure"] == "PCCM_ModelCheckRefusalSubject"
    assert plan.reading_formula(entry) == "=PCCM_ModelCheckRefusalSubject()"
    check = next(c for c in plan.ordered_checks if c["check_id"] == "CAL-010")
    assert str(check["subject"]) == "{calculation_refusal_subject}", check["subject"]
    assert str(check["message"]) == "{calculation_refusal_detail}", check["message"]
    adapter = _procedure("modResultsState", "PCCM_ModelCheckRefusalSubject")
    assert "Application.Volatile True" in adapter
    assert "modCalcReport.CalcReportDerivedStatus(detail, subject)" in adapter
    assert "CVErr(xlErrValue)" in adapter
    for banned in ("PCCM_CalculationStatus", "PCCM_CalculationAttemptDetail", "Static"):
        assert banned not in adapter, f"the subject adapter reaches {banned}"
    # ONE PREPARATION, TWO OUTPUTS. There is no second traversal to find the id.
    assert adapter.count("CalcReportDerivedStatus") == 1


# --- the source mutations -------------------------------------------------
def _subject_sources(text: str) -> list[str]:
    """The right-hand side of every subject assignment, compound lines included."""
    return [statement[len("subject = "):].strip()
            for statement in _subject_statements(text)]


def _assert_sources_are_permanent_ids(text: str) -> None:
    """Every subject is a permanent id, a clear, or a local that is only ever
    one of those.

    THE INDIRECTION IS ALLOWED AND CHECKED, NOT WAVED THROUGH. AccumulateTotals
    already reads `who = audits(index).PermanentId` for the sentence its
    delegate writes, and P9-3 hands that same local out rather than reading the
    array a second time. So a bare local passes only if EVERY assignment to it
    in the module ends in `.PermanentId`: one that is ever set from a currency,
    a row number or a literal fails here exactly as a direct one would.
    """
    import re

    for source in _subject_sources(text):
        if source == "vbNullString" or source.endswith(".PermanentId"):
            continue
        assert re.fullmatch(r"[A-Za-z_]\w*", source), (
            f"subject is set from something that is not a permanent id: {source}")
        binds = [statement[len(source) + 3:].strip()
                 for line in text.splitlines() for statement in
                 (part.strip() for part in line.split(":"))
                 if statement.startswith(f"{source} = ")]
        assert binds, f"subject is set from {source}, which is never assigned"
        for bind in binds:
            assert bind.endswith(".PermanentId"), (
                f"subject is set from {source}, which is assigned {bind}")


# EVERY THREADED OWNER, NOT JUST THE FIRST ONE. P9-2B proved the rule on
# modCalcCheck; P9-3 adds four more procedures across two modules, and a rule
# demonstrated on one of six is not a rule.
_ID_SITES = [
    ("modCalcCheck", "subject = model.Drivers(LBound(model.Drivers) + index).PermanentId",
     "model.Drivers(LBound(model.Drivers) + index)"),
    ("modCalcResolve", "subject = driver.PermanentId", "driver"),
    ("modCalcAnalytical", "subject = drivers(index).PermanentId", "drivers(index)"),
    ("modCalcAnalytical",
     "subject = drivers(LBound(drivers) + order(slot)).PermanentId",
     "drivers(LBound(drivers) + order(slot))"),
    ("modCalcReport", "subject = package.Model.Drivers(index).PermanentId",
     "package.Model.Drivers(index)"),
    ("modCalcReport", "subject = package.Drivers(index).PermanentId",
     "package.Drivers(index)"),
]


@pytest.mark.parametrize("replacement,label", [
    (".Currency", "a description-like field"),
    ("CStr(rowIndex)", "a worksheet row number"),
    ('"CL-0001"', "a hard-coded id"),
])
@pytest.mark.parametrize("module,statement,owner", _ID_SITES,
                         ids=[f"{m}:{o[:24]}" for m, _s, o in _ID_SITES])
def test_46o_mutation_the_subject_stops_being_a_permanent_id(
        module: str, statement: str, owner: str,
        replacement: str, label: str) -> None:
    """A subject taken from anything but the driver's permanent id - its
    currency, its row number, a literal - reads plausibly and is wrong."""
    text = _code(module)
    assert statement in text, (module, statement)
    mutated = text.replace(
        statement,
        f"subject = {owner}{replacement}" if replacement.startswith(".")
        else f"subject = {replacement}", 1)
    assert mutated != text, label
    with pytest.raises(AssertionError):
        _assert_sources_are_permanent_ids(mutated)
    _assert_sources_are_permanent_ids(text)


def test_46o1_mutation_the_local_the_subject_borrows_stops_being_an_id() -> None:
    """THE INDIRECTION CANNOT BE THE WAY ROUND THE RULE. AccumulateTotals hands
    out `who`, the local its delegate's sentence is built from. If `who` were
    ever assigned something that is not a permanent id, the subject would stop
    being one without a single subject line changing."""
    text = _code("modCalcAnalytical")
    assert "subject = who" in text
    mutated = text.replace("who = audits(index).PermanentId",
                           "who = audits(index).Description", 1)
    assert mutated != text
    with pytest.raises(AssertionError):
        _assert_sources_are_permanent_ids(mutated)
    _assert_sources_are_permanent_ids(text)


def test_46p_mutation_a_driver_specific_refusal_leaves_the_subject_blank() -> None:
    """The set removed from the loop: every per-driver refusal in modCalcCheck
    would then report no id at all."""
    text = _code("modCalcCheck")
    mutated = text.replace(
        "        subject = model.Drivers(LBound(model.Drivers) + index).PermanentId\n", "", 1)
    assert mutated != text
    assert not [line for line in mutated.splitlines()
                if line.strip().startswith("subject = ")
                and line.strip().endswith(".PermanentId")], "the mutation changed nothing"
    # WHICH test_46j REFUSES, because a threaded owner must actually thread.
    assert [line for line in text.splitlines()
            if line.strip().startswith("subject = ")
            and line.strip().endswith(".PermanentId")]


def test_46q_mutation_a_stale_subject_survives_a_success() -> None:
    """The clear removed: a subject set for the last driver would attach itself
    to the next model-wide refusal, naming a driver that had nothing to do
    with it."""
    text = _code("modCalcCheck")
    mutated = text.replace("    subject = vbNullString\n    CheckResolvedModel = True",
                           "    CheckResolvedModel = True", 1)
    assert mutated != text
    lines = [line.strip() for line in mutated.splitlines()]
    sets = [i for i, line in enumerate(lines)
            if line.startswith("subject = ") and not line.endswith("vbNullString")]
    clears = [i for i, line in enumerate(lines) if line == "subject = vbNullString"]
    success = [i for i, line in enumerate(lines) if line == "CheckResolvedModel = True"]
    assert not any(max(sets) < clear < success[-1] for clear in clears), (
        "the mutation changed nothing")


def test_46r_mutation_the_subject_is_taken_from_the_persisted_attempt() -> None:
    def edit(block):
        for check in block["checks"]:
            if check["check_id"] == "CAL-010":
                check["subject"] = "{calculation_attempt_detail}"
    plan = _mutate_manifest(edit)
    broken = _evaluate(plan, {"calculation_state": "INVALID",
                              "calculation_refusal_subject": "CL-0001",
                              "calculation_attempt_detail": _refusal("CL-0009")})
    row = next(r for r in broken["shown"] if r["check_id"] == "CAL-010")
    assert "CL-0009" in str(row["subject"]), "the mutation changed nothing"
    with pytest.raises(AssertionError):
        _assert_live_subject(plan)
    _assert_live_subject(_plan())


def _assert_live_subject(plan: ModelCheckPlan) -> None:
    check = next(c for c in plan.ordered_checks if c["check_id"] == "CAL-010")
    assert str(check["subject"]) == "{calculation_refusal_subject}", check["subject"]


def test_46t_the_plumbing_reversal_hides_nothing() -> None:
    """THE REVERSAL IS WHY NO HISTORICAL DIGEST MOVED, so it has to be proved it
    cannot absorb anything else. A changed condition, a reworded message, a
    flipped Boolean and a moved constant all survive it - and then the digests
    that stayed put fail, exactly as they always did."""
    import hashlib
    import sys as _sys

    _sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_subject_plumbing import reverse_subject_plumbing

    text = (SRC / "modCalcCheck.bas").read_text(encoding="utf-8")
    baseline = subprocess.run(["git", "show", "ad78988:pccm/src/vba/modCalcCheck.bas"],
                              cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE,
                              text=True).stdout
    # AS IT STANDS, REVERSING RESTORES THE PHASE-7 BYTES EXACTLY.
    assert reverse_subject_plumbing("modCalcCheck", text) == baseline, (
        "the subject plumbing is not the only thing that changed in modCalcCheck")

    semantic = (
        ("a validation condition", "driver.MinValue > driver.MostLikely",
         "driver.MinValue >= driver.MostLikely"),
        ("a Boolean outcome", "CheckResolvedModel = True", "CheckResolvedModel = False"),
        ("a refusal message", "Quantity must be strictly positive",
         "Quantity must be positive"),
        ("a constant", "PROFILE_SUM_TARGET", "PROFILE_SUM_LIMIT"),
    )
    for label, before, after in semantic:
        mutated = text.replace(before, after, 1)
        assert mutated != text, label
        restored = reverse_subject_plumbing("modCalcCheck", mutated)
        assert restored != baseline, (
            f"the reversal absorbed {label}; it would have hidden a real change")
        assert hashlib.sha256(restored.encode()).hexdigest() != \
            hashlib.sha256(baseline.encode()).hexdigest()


@pytest.mark.parametrize("module,semantic", [
    ("modCalcAnalytical", ("an identity tolerance", "TOL_PROFILING_SUM_ABSOLUTE",
                           "TOL_IDENTITY_RELATIVE_COEFFICIENT")),
    ("modCalcAnalytical", ("an arithmetic expression", "10 + slot", "11 + slot")),
    ("modCalcAnalytical", ("a Boolean outcome", "Reconcile = True", "Reconcile = False")),
    ("modCalcAnalytical", ("a refusal message", "canonical driver order",
                           "canonical order")),
    ("modCalcReport", ("a refusal message",
                       "the inflation profile is not in the resolved reference set",
                       "the inflation profile is unknown")),
    ("modCalcReport", ("a Boolean outcome", "BuildAudits = True", "BuildAudits = False")),
    ("modCalcReport", ("a call the preparation makes", "CountCurrencyReferences package",
                       "CountCurrencyReferences package.Model")),
], ids=lambda v: v if isinstance(v, str) else v[0])
def test_46t1_the_reversal_hides_nothing_in_the_two_owners_p9_3_reached(
        module: str, semantic: tuple) -> None:
    """THE SAME PROOF, FOR THE MODULES P9-3 TOUCHED. Their pins - the pre-Run-7
    digest for modCalcAnalytical, the accepted reporter prefix for modCalcReport
    - did not move because the reversal restores them. That claim is only worth
    anything if the reversal cannot swallow a real change as well."""
    import sys as _sys

    _sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_subject_plumbing import reverse_subject_plumbing

    label, before, after = semantic
    text = (SRC / f"{module}.bas").read_text(encoding="utf-8")
    baseline = subprocess.run(["git", "show", f"ad78988:pccm/src/vba/{module}.bas"],
                              cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE,
                              text=True).stdout
    if module == "modCalcReport":
        banner = ("' ==========================================================================\n"
                  "' STEP 11 ADDITION - THE PHASE-6 PREPARATION BRIDGE\n")
        text, baseline = text[:text.index(banner)], baseline[:baseline.index(banner)]
    assert reverse_subject_plumbing(module, text) == baseline, (
        f"the subject plumbing is not the only thing that changed in {module}")
    mutated = text.replace(before, after, 1)
    assert mutated != text, (module, label)
    assert reverse_subject_plumbing(module, mutated) != baseline, (
        f"the reversal absorbed {label} in {module}; it would have hidden a real change")


def test_46u_the_reversal_refuses_to_run_on_text_it_does_not_recognise() -> None:
    """A comment block that is not what the reversal removes means the module
    changed in a way nobody described. It refuses rather than guessing."""
    import sys as _sys

    _sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_subject_plumbing import COMMENT_ADDITIONS, reverse_subject_plumbing

    text = (SRC / "modCalcResolve.bas").read_text(encoding="utf-8")
    block = COMMENT_ADDITIONS["modCalcResolve"][0]
    with pytest.raises(AssertionError, match="not the text this reversal removes"):
        reverse_subject_plumbing("modCalcResolve", text.replace(block, "", 1))
    # AND A MODULE IT KNOWS NOTHING ABOUT COMES BACK UNTOUCHED.
    assert reverse_subject_plumbing("modSimEngine", "anything") == "anything"


def test_46s_the_historical_byte_pins_were_not_overwritten() -> None:
    """§4, AND IT IS HERE BECAUSE THE HAZARD ALREADY HAPPENED ONCE. A regex that
    took the first match repointed the Run-6 CLOSURE digest instead of the
    current pin. History is history: it is verified against the commit it is
    about, never against today's tree."""
    import hashlib
    import re

    source = (PCCM_ROOT / "tests" / "test_phase6_integration_source.py").read_text(
        encoding="utf-8")
    closure = re.search(r'STEP13_CLOSURE_COMMIT = "([0-9a-f]{40})"', source).group(1)
    frozen = source[source.index("FROZEN_SOURCE = {"):]
    frozen = frozen[:frozen.index("}")]
    pins = dict(re.findall(r'"(\w+)": "([0-9a-f]{64})"', frozen))
    assert pins, "the historical pin table is empty"
    for module, digest in pins.items():
        blob = subprocess.run(["git", "show", f"{closure}:pccm/src/vba/{module}.bas"],
                              cwd=REPO_ROOT, capture_output=True)
        assert blob.returncode == 0, module
        actual = hashlib.sha256(blob.stdout).hexdigest()
        assert actual == digest, (
            f"the historical pin for {module} no longer matches the bytes at "
            f"{closure[:7]}; a current digest has been written over history")
    # AND THE MODULES THIS PHASE TOUCHED ARE AMONG THEM, so the check is not
    # about somebody else's files.
    assert {"modCalcReport"} <= set(pins), sorted(pins)


# ===========================================================================
# F. WORKSHEET SAFETY
# ===========================================================================
_PERSISTING = {
    "PCCM_CalculationStatus": "writes _Calc C19:C20 through WriteStatusBlock",
    "PCCM_SimulationStatus": "writes _SimData D28:D29 through WriteStatusBlock",
}

_COMMAND_ENTRY_POINTS = ("PCCM_Calculate", "PCCM_RunSimulation", "PCCM_RunSensitivity",
                         "PCCM_RunAnnualStochastic", "PCCM_ApplyTimeline",
                         "PCCM_AddCostLine", "PCCM_AddRisk", "PCCM_DeleteCostLine",
                         "PCCM_DeleteRisk", "PCCM_AutomationBegin", "PCCM_AutomationEnd")

_WRITE_TOKENS = (".Value2 =", ".Value =", ".Formula =", "ClearContents", ".Delete",
                 "ListRows.Add", "Application.ScreenUpdating", "Application.Calculation",
                 "Application.EnableEvents")


def _module(name: str) -> str:
    return (SRC / f"{name}.bas").read_text(encoding="utf-8")


def _code(name: str) -> str:
    return "\n".join(line for line in _module(name).splitlines()
                     if not line.lstrip().startswith("'"))


def _procedure(module: str, name: str) -> str:
    """One procedure's body, comment lines stripped, bounded at its own End."""
    lines = _code(module).splitlines()
    starts = [i for i, line in enumerate(lines)
              if line.startswith(("Public Function " + name + "(",
                                  "Private Function " + name + "(",
                                  "Public Sub " + name + "(",
                                  "Private Sub " + name + "("))]
    assert len(starts) == 1, f"{module}.{name} is declared {len(starts)} times"
    start = starts[0]
    for offset in range(start + 1, len(lines)):
        if lines[offset].startswith(("End Function", "End Sub")):
            return "\n".join(lines[start:offset + 1])
    raise AssertionError(f"{module}.{name} has no end")


def test_41_every_cell_called_procedure_is_read_only() -> None:
    """THE CLAIM THE WHOLE SHEET RESTS ON, checked over the real call graph and
    not over a list somebody typed."""
    plan = _plan()
    called = sorted({str(e["procedure"]) for e in plan.readings["rows"]
                     if e["kind"] == "procedure"})
    assert called, "no cell-called procedure was found; the control proves nothing"
    owners = {name: module for module in
              ("modResultsState", "modStructuralCheck", "modCalcReport", "modSimReport",
               "modSimAnnualStore")
              for name in _public_names(module)}
    seen: set[str] = set()

    def walk(name: str, trail: list[str]) -> None:
        if name in seen:
            return
        seen.add(name)
        module = owners.get(name)
        assert module is not None, f"{' -> '.join(trail)}: nothing owns {name}"
        body = _procedure(module, name)
        for token in _WRITE_TOKENS:
            assert token not in body, (
                f"{' -> '.join(trail + [name])} contains {token!r}: a cell-called path "
                "may not change the workbook")
        for persisting in _PERSISTING:
            assert persisting not in body, (
                f"{' -> '.join(trail + [name])} reaches {persisting}, which "
                f"{_PERSISTING[persisting]}")
        for command in _COMMAND_ENTRY_POINTS:
            assert command not in body, (
                f"{' -> '.join(trail + [name])} invokes the command endpoint {command}")
        for callee in owners:
            if callee != name and (f"{callee}(" in body or f".{callee}(" in body):
                walk(callee, trail + [name])
        for private in _private_names(module):
            if private != name and f"{private}(" in body:
                walk_private(module, private, trail + [name])

    def walk_private(module: str, name: str, trail: list[str]) -> None:
        key = f"{module}.{name}"
        if key in seen:
            return
        seen.add(key)
        body = _procedure(module, name)
        for token in _WRITE_TOKENS:
            assert token not in body, (
                f"{' -> '.join(trail + [key])} contains {token!r}")
        for persisting in _PERSISTING:
            assert persisting not in body, f"{' -> '.join(trail + [key])} reaches {persisting}"
        for private in _private_names(module):
            if private != name and f"{private}(" in body:
                walk_private(module, private, trail + [key])

    for procedure in called:
        walk(procedure, [])


def _public_names(module: str) -> list[str]:
    import re

    return re.findall(r"^Public (?:Function|Sub) (\w+)", _code(module), re.MULTILINE)


def _private_names(module: str) -> list[str]:
    import re

    return re.findall(r"^Private (?:Function|Sub) (\w+)", _code(module), re.MULTILINE)


def test_41b_the_sensitivity_presentation_reaches_no_vba_at_all() -> None:
    """FORMULA-ONLY, AND MIRRORED WHOLE. The availability sentence is the
    Sensitivity sheet's own worksheet formula over `_SimData`; Model Check copies
    the sentence and adds no judgement of its own to it."""
    plan = _plan()
    entry = next(e for e in plan.readings["rows"] if e["key"] == "sensitivity_availability")
    assert entry["kind"] == "sensitivity_availability"
    formula = plan.reading_formula(entry)
    assert "PCCM_" not in formula, f"the sensitivity reading calls VBA: {formula}"
    sensitivity = yaml.safe_load(MANIFEST.read_text(encoding="utf-8"))[
        "phase6_shell"]["sensitivity"]
    column = str(sensitivity["columns"][1]["column"])
    assert f"{sensitivity['sheet']}!${column}${sensitivity['availability_row']}" in formula
    # AND THE ROW THAT SHOWS IT IS INFO, because a persisted comparison is
    # context: it is blind to model drift and is paired with the live simulation
    # state rather than judged on its own.
    check = next(c for c in plan.ordered_checks if c["check_id"] == "SEN-010")
    assert check["severity"] == plan.informational


def test_41c_the_annual_readings_are_the_accepted_phase_8_adapters() -> None:
    """PHASE 9 ADDED ONE ADAPTER AND BORROWED FOUR. The annual read path is the
    one P8-1 accepted; nothing here reaches past it into modSimAnnualStore."""
    plan = _plan()
    annual = [str(e["procedure"]) for e in plan.readings["rows"]
              if e["kind"] == "procedure" and str(e["key"]).startswith("annual_")]
    assert annual == ["PCCM_ResultsAnnualDistributionState",
                      "PCCM_ResultsAnnualProfileState",
                      "PCCM_ResultsAnnualProfilePx",
                      "PCCM_ResultsAnnualYearCount"], annual
    joined = "\n".join(str(v) for v in _cells(plan).values())
    for owner in ("PCCM_AnnualDistributionState()", "PCCM_AnnualProfileState()"):
        assert owner not in joined, f"a Model Check cell reaches past the adapter to {owner}"


def test_42_the_persisting_entry_points_are_not_reachable_from_a_cell() -> None:
    plan = _plan()
    formulas = [plan.reading_formula(e) for e in plan.readings["rows"]]
    for cells in _cells(plan).values():
        if isinstance(cells, str) and cells.startswith("="):
            formulas.append(cells)
    joined = "\n".join(formulas)
    for persisting in _PERSISTING:
        assert persisting not in joined, f"a Model Check cell calls {persisting}"


def test_43_the_structural_path_is_still_read_only() -> None:
    """`modStructuralCheck` REPORTS; it never repairs. Checked module-wide, not
    only along the path this sheet happens to take today."""
    code = _code("modStructuralCheck")
    for token in _WRITE_TOKENS:
        assert token not in code, f"modStructuralCheck contains {token!r}"
    assert "PCCM_StructuralReport" in code
    assert _procedure("modStructuralCheck", "PCCM_StructuralReport").count("\n") <= 3


def test_44_the_new_calculation_adapter_is_the_authorised_split() -> None:
    calc = _code("modCalcReport")
    # THE OUT-PARAMETER IS REQUIRED, NOT OPTIONAL. A typed Optional with no
    # default is a VBA COMPILE ERROR, and both callers supply the argument
    # anyway - the same settlement modDrivers.HighestIssued reached after it
    # cost a Gate-B build.
    assert ("Public Function CalcReportDerivedStatus(ByRef detail As String, "
            "ByRef subject As String) As String") in calc
    body = _procedure("modCalcReport", "CalcReportDerivedStatus")
    assert "DeriveStatus(" in body and "PrepareCurrentCalculation(" in body
    assert "WriteStatusBlock" not in body
    # AND THE PERSISTING ENTRY POINT IS UNTOUCHED: it still derives AND writes.
    status = _procedure("modCalcReport", "PCCM_CalculationStatus")
    assert "WriteStatusBlock status" in status

    state = _code("modResultsState")
    assert "Public Function PCCM_ModelCheckCalculationState() As Variant" in state
    adapter = _procedure("modResultsState", "PCCM_ModelCheckCalculationState")
    assert "Application.Volatile True" in adapter
    assert "modCalcReport.CalcReportDerivedStatus(detail, subject)" in adapter
    assert "CVErr(xlErrValue)" in adapter, "the adapter must fail loud, not wrong"
    assert "PCCM_CalculationStatus" not in adapter


def test_45_the_adapter_invents_no_state_word() -> None:
    """Phase 9 owns no vocabulary. The four calculation words arrive spelled the
    way modCalcReport spells them, and modResultsState defines none of them."""
    adapter = _procedure("modResultsState", "PCCM_ModelCheckCalculationState")
    for word in ("NOT CALCULATED", "CURRENT", "STALE", "INVALID", "PASS", "WARNING",
                 "ERROR"):
        assert f'"{word}"' not in adapter, f"the adapter spells {word!r}"


# THE PHASE-9 PRODUCTION CORRECTIONS, DECLARED
# ---------------------------------------------------------------------------
# P9-2 and P9-2A were purely ADDITIVE and the control said so. P9-2B is not:
# threading a structured subject through the preparation edits signatures and
# call sites in modules the Phase-7 evidence ran against, and it was authorised
# on exactly those terms.
#
# SO THE CLAIM BECOMES MECHANICAL RATHER THAN BLANKET. A declared file may lose
# a line ONLY if the same line comes back with nothing added but the subject
# plumbing. That is stronger than a hand-listed set of permitted removals: it
# proves, line by line, that no condition, no message, no Boolean and no
# arithmetic moved - because if any of them had, the normalised line would not
# match.
# AND THE DEFINITION OF "PLUMBING" IS THE SHARED ONE. This control used to keep
# its own token list, which meant two descriptions of the same rule and two
# chances for one of them to drift. It now asks tests/vba_subject_plumbing.py -
# the same function the byte-for-byte module reversals ask - so a form that
# reverses cleanly there and a form that normalises cleanly here are the same
# form, by construction.
DECLARED_PHASE9_CORRECTIONS = {
    "pccm/src/vba/modResultsState.bas": "additive",
    "pccm/src/vba/modCalcReport.bas": "plumbing",
    "pccm/src/vba/modCalcCheck.bas": "plumbing",
    "pccm/src/vba/modCalcResolve.bas": "plumbing",
    # P9-3 completes the coverage in the last owner that held a permanent id it
    # was not handing out.
    "pccm/src/vba/modCalcAnalytical.bas": "plumbing",
}


def _normalise(line: str) -> str:
    """A line with the subject plumbing taken back out of it.

    The rule is not restated here: reverse_line is the SAME definition the
    module-wide byte reversals use, so a line that reverses to the accepted
    bytes and a line that normalises to its undamaged self cannot disagree.
    Whitespace is then collapsed, because a line that gained a parameter may
    have been re-wrapped and the statement is what is being compared.
    """
    from vba_subject_plumbing import reverse_line

    return " ".join((reverse_line(line) or "").split())


def _logical(lines: list[str]) -> list[str]:
    """VBA continuation lines rejoined.

    A parameter added to a signature can push it onto another line, so one
    REMOVED physical line becomes two INSERTED ones. Comparing physical lines
    would call that a rewrite; comparing the statements they belong to sees it
    for what it is.
    """
    joined, buffer = [], ""
    for line in lines:
        buffer = (buffer + " " + line.strip()) if buffer else line.rstrip()
        if buffer.rstrip().endswith(" _"):
            buffer = buffer.rstrip()[:-1]
            continue
        joined.append(buffer)
        buffer = ""
    if buffer:
        joined.append(buffer)
    return joined


def test_46_every_production_change_is_declared_and_is_only_plumbing() -> None:
    head = "ad78988"
    changed = subprocess.run(["git", "diff", "--name-status", head, "--", "pccm/src"],
                             cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE,
                             text=True).stdout
    modified, added = [], []
    for line in changed.splitlines():
        if not line.strip():
            continue
        state, path = line.split("\t", 1)[0].strip(), line.split("\t", 1)[1].strip()
        assert not state.startswith("D"), f"a module the evidence ran against was removed: {path}"
        (added if state.startswith("A") else modified).append(path)

    phase8 = {"pccm/src/vba/modSimAnnualStore.bas", "pccm/src/vba/modSimPostReport.bas",
              "pccm/src/vba/modSimReport.bas"}
    undeclared = [p for p in modified
                  if p not in DECLARED_PHASE9_CORRECTIONS and p not in phase8]
    assert not undeclared, f"production changed without being declared: {undeclared}"

    for path, kind in DECLARED_PHASE9_CORRECTIONS.items():
        diff = subprocess.run(["git", "diff", head, "--", path], cwd=REPO_ROOT,
                              check=True, stdout=subprocess.PIPE, text=True).stdout
        removed = [line[1:] for line in diff.splitlines()
                   if line.startswith("-") and not line.startswith("---")]
        inserted = [line[1:] for line in diff.splitlines()
                    if line.startswith("+") and not line.startswith("+++")]
        if kind == "additive":
            assert not removed, f"{path} is declared additive and removes {len(removed)} line(s)"
            continue
        # EVERY REMOVED LINE COMES BACK, PLUMBING APART.
        available = [_normalise(line) for line in _logical(inserted)]
        for line in _logical(removed):
            wanted = _normalise(line)
            assert wanted in available, (
                f"{path} lost a line that is not merely re-plumbed:\n  {line.strip()}")
            available.remove(wanted)


def test_46_1_the_plumbing_normaliser_is_not_a_blanket_pass() -> None:
    """SO THE CONTROL ABOVE MEANS SOMETHING. Taking the subject tokens out must
    not take a condition, a message or a Boolean out with them."""
    assert _normalise("If x > 0 Then") == "If x > 0 Then"
    assert _normalise('detail = "cost line " & id') == 'detail = "cost line " & id'
    assert _normalise("CheckResolvedModel = True") == "CheckResolvedModel = True"
    # A REAL EDIT SURVIVES NORMALISATION AND SO IS STILL CAUGHT.
    assert _normalise("If x > 0 Then") != _normalise("If x >= 0 Then")
    assert _normalise("    detail = vbNullString: subject = vbNullString") == \
        _normalise("    detail = vbNullString")
    assert _normalise("ByRef detail As String, ByRef subject As String") == \
        _normalise("ByRef detail As String")


def test_46_2_no_module_changed_its_line_endings() -> None:
    """The P9-2 defect, still refused - and now over four modules rather than
    two."""
    head = "ad78988"
    for path in DECLARED_PHASE9_CORRECTIONS:
        before = subprocess.run(["git", "show", f"{head}:{path}"], cwd=REPO_ROOT,
                                capture_output=True)
        if before.returncode != 0:
            continue
        raw = (REPO_ROOT / path).read_bytes()
        was_crlf, is_crlf = b"\r\n" in before.stdout, b"\r\n" in raw
        assert was_crlf == is_crlf, f"{path} changed line-ending convention"
        if is_crlf:
            assert raw.count(b"\n") == raw.count(b"\r\n"), f"{path} mixes line endings"


# ===========================================================================
# G. MUTATIONS - ZERO SURVIVORS
# ===========================================================================
# Each entry breaks one thing the sheet promises, and the control that owns that
# promise must refuse the broken build. A mutation nothing catches is a promise
# nothing checks.

def _mutate_manifest(edit) -> ModelCheckPlan:
    manifest = _manifest()
    edit(manifest["phase9_shell"]["model_check"])
    return _plan(manifest)


def test_47_mutation_the_advisory_boundary_becomes_inclusive() -> None:
    def edit(block):
        for check in block["checks"]:
            check["condition"] = check["condition"].replace(
                "{requested_iterations}<{recommended_iterations}",
                "{requested_iterations}<={recommended_iterations}")
    plan = _mutate_manifest(edit)
    result = _evaluate(plan, {"calculation_state": "CURRENT",
                              "requested_iterations": plan.recommended_iterations})
    assert "INP-010" in _ids(result), (
        "the boundary mutation changed nothing; the control below proves nothing")
    with pytest.raises(AssertionError):
        _assert_boundary(plan)


def _assert_boundary(plan: ModelCheckPlan) -> None:
    at = _evaluate(plan, {"calculation_state": "CURRENT",
                          "requested_iterations": plan.recommended_iterations})
    assert "INP-010" not in _ids(at)


def test_48_mutation_the_threshold_drifts_from_the_contract() -> None:
    plan = _plan()
    for drift in (-1, +1):
        broken = copy.deepcopy(_projection())
        broken["advisory"]["threshold"] = plan.recommended_iterations + drift
        with pytest.raises(ValueError, match="threshold"):
            validate_phase9_inspection(broken)


def test_49_mutation_not_calculated_demoted_to_info() -> None:
    def edit(block):
        for check in block["checks"]:
            if check["check_id"] == "CAL-020":
                check["severity"] = "INFO"
    plan = _mutate_manifest(edit)
    result = _evaluate(plan, {"calculation_state": "NOT CALCULATED"})
    assert _counts(result) == (0, 0)
    assert result["summary"]["overall_status"] == "PASS"
    # WHICH IS EXACTLY WHAT SCENARIO A REFUSES.
    assert _counts(_evaluate(_plan(), {"calculation_state": "NOT CALCULATED"})) == (0, 1)


def test_50_mutation_an_optional_publication_promoted_to_warning() -> None:
    def edit(block):
        for check in block["checks"]:
            if check["check_id"] == "SIM-050":
                check["severity"] = "WARNING"
                # THE FLAG COMES OFF TOO, because the loader now refuses the
                # flagged form outright - test_38c is that refusal. Removing it
                # is what lets this mutation reach the SHEET, so the runtime
                # consequence is demonstrated as well as the build-time one.
                check.pop("optional_publication", None)
    plan = _mutate_manifest(edit)
    result = _evaluate(plan, {"calculation_state": "CURRENT"})
    assert _counts(result) == (0, 1), "the mutation changed nothing"
    assert result["summary"]["overall_status"] == "WARNING"
    assert _evaluate(_plan(), {"calculation_state": "CURRENT"}
                     )["summary"]["overall_status"] == "PASS"


def test_51_mutation_sim_invalid_double_counted() -> None:
    def edit(block):
        for check in block["checks"]:
            if check["check_id"] == "SIM-020":
                check["severity"] = "ERROR"
    plan = _mutate_manifest(edit)
    result = _evaluate(plan, {"calculation_state": "NOT CALCULATED",
                              "simulation_state": "INVALID"})
    assert _counts(result) == (1, 1), "the mutation changed nothing"
    assert _counts(_evaluate(_plan(), {"calculation_state": "NOT CALCULATED",
                                       "simulation_state": "INVALID"})) == (0, 1)


def test_52_mutation_info_allowed_to_alter_the_overall_status() -> None:
    plan = _plan()
    cells = _cells(plan)
    address = f"{plan.value_column}{plan.summary_row('overall_status')}"
    cells[address] = (f'=IF({plan.summary_cell("error_count")}>0,"ERROR",'
                      f'IF({plan.summary_cell("warning_count")}+'
                      f'{plan.summary_cell("info_count")}>0,"WARNING","PASS"))')
    broken = _evaluate(plan, {"calculation_state": "CURRENT"}, cells=cells)
    assert broken["summary"]["overall_status"] == "WARNING", "the mutation changed nothing"
    healthy = _evaluate(plan, {"calculation_state": "CURRENT"})
    assert healthy["summary"]["overall_status"] == "PASS"
    assert plan.summary_cell("info_count") not in plan.summary_formula("overall_status")


def test_53_mutation_the_overflow_total_counts_only_the_window() -> None:
    plan = _plan()
    cells = _cells(plan)
    address = f"{plan.value_column}{plan.summary_row('error_count')}"
    surplus = f"+{plan.reading_cell('structural_surplus')}"
    assert surplus in cells[address], cells[address]
    cells[address] = cells[address].replace(surplus, "")
    faults = plan.row_window + 25
    broken = _evaluate(plan, {"structural_report": _faults(faults),
                              "calculation_state": "INVALID"}, cells=cells)
    healthy = _evaluate(plan, {"structural_report": _faults(faults),
                               "calculation_state": "INVALID"})
    assert int(broken["summary"]["error_count"]) < int(healthy["summary"]["error_count"])
    assert int(healthy["summary"]["error_count"]) >= faults
    assert int(broken["summary"]["total_checks"]) < faults, (
        "the mutation left the total honest; the control proves nothing")


def test_54_mutation_deduplication_by_message() -> None:
    plan = _plan()
    cells = _cells(plan)
    subject_column = plan.candidate_column("subject")
    message_column = plan.candidate_column("message")
    for row in plan.declared_rows():
        address = f"{plan.candidate_column('dedup_key')}{row}"
        cells[address] = cells[address].replace(f"${subject_column}{row}",
                                                f"${message_column}{row}")
    manifest = _manifest()
    checks = manifest["phase9_shell"]["model_check"]["checks"]
    twin = copy.deepcopy(next(c for c in checks if c["check_id"] == "CAL-020"))
    twin["subject"] = "A different subject"
    checks.append(twin)
    mutated_plan = _plan(manifest)
    mutated_cells = _cells(mutated_plan)
    for row in mutated_plan.declared_rows():
        address = f"{mutated_plan.candidate_column('dedup_key')}{row}"
        mutated_cells[address] = mutated_cells[address].replace(
            f"${subject_column}{row}", f"${message_column}{row}")
    broken = _evaluate(mutated_plan, {"calculation_state": "NOT CALCULATED"},
                       cells=mutated_cells)
    healthy = _evaluate(mutated_plan, {"calculation_state": "NOT CALCULATED"})
    assert _ids(broken).count("CAL-020") == 1, "the mutation changed nothing"
    assert _ids(healthy).count("CAL-020") == 2
    assert cells is not None


def test_55_mutation_an_unused_slot_rendered_as_a_real_row() -> None:
    plan = _plan()
    cells = _cells(plan)
    for row in plan.register_rows():
        for column in plan.register["columns"]:
            address = f"{column['column']}{row}"
            cells[address] = cells[address].replace("NA()", '""')
    broken = _evaluate(plan, {"calculation_state": "CURRENT"}, cells=cells)
    assert not broken["unused"], "the mutation changed nothing"
    assert any(row["check_id"] == "" for row in broken["register"])
    healthy = _evaluate(plan, {"calculation_state": "CURRENT"})
    assert healthy["unused"]


def test_56_mutation_a_write_capable_function_in_a_cell() -> None:
    def edit(block):
        for entry in block["evaluation"]["readings"]["rows"]:
            if entry["key"] == "calculation_state":
                entry["procedure"] = "PCCM_CalculationStatus"
    plan = _mutate_manifest(edit)
    joined = "\n".join(str(v) for v in _cells(plan).values())
    assert "PCCM_CalculationStatus" in joined, "the mutation changed nothing"
    inspection = build_phase9_inspection(plan)
    with pytest.raises(ValueError, match="persists"):
        validate_phase9_inspection(inspection)


def test_57_mutation_the_persisted_status_used_instead_of_the_live_adapter() -> None:
    def edit(block):
        for entry in block["evaluation"]["readings"]["rows"]:
            if entry["key"] == "calculation_state":
                entry.clear()
                entry.update({"key": "calculation_state", "row": 125,
                              "label": "Calculation State", "kind": "mirror",
                              "source_block": "run_stamp",
                              "source_key": "simulation_status", "format": "text"})
    plan = _mutate_manifest(edit)
    formula = plan.reading_formula(
        next(e for e in plan.readings["rows"] if e["key"] == "calculation_state"))
    assert "PCCM_ModelCheckCalculationState" not in formula, "the mutation changed nothing"
    with pytest.raises(AssertionError):
        _assert_live_adapter(plan)


def _assert_live_adapter(plan: ModelCheckPlan) -> None:
    entry = next(e for e in plan.readings["rows"] if e["key"] == "calculation_state")
    assert plan.reading_formula(entry) == "=PCCM_ModelCheckCalculationState()"


def test_58_mutation_the_order_becomes_discovery_order() -> None:
    """A register that took the manifest's listing order would be a register
    whose order nobody owns. The plan sorts; removing the sort changes the
    answer, and test_24 is what refuses it."""
    plan = _plan()
    manifest = _manifest()
    manifest["phase9_shell"]["model_check"]["checks"].reverse()
    reversed_declaration = manifest["phase9_shell"]["model_check"]["checks"]
    unsorted = [str(c["check_id"]) for c in reversed_declaration]
    sorted_ids = [str(c["check_id"]) for c in _plan(manifest).ordered_checks]
    assert unsorted != sorted_ids, "the register order is already the declaration order"
    assert sorted_ids == [str(c["check_id"]) for c in plan.ordered_checks]


def test_59_mutation_a_condition_names_a_reading_nobody_publishes() -> None:
    from pccm_builder.spec_loader import SpecError

    def edit(block):
        block["checks"][0]["condition"] = '{no_such_reading}="X"'
    with pytest.raises(SpecError, match="no_such_reading"):
        _mutate_manifest(edit)


def test_60_mutation_the_structural_slots_shrink_below_the_window() -> None:
    from pccm_builder.spec_loader import SpecError

    def edit(block):
        block["evaluation"]["candidates"]["structural_slots"] = block["row_window"] - 1
    with pytest.raises(SpecError, match="structural_slots"):
        _mutate_manifest(edit)


def test_61_mutation_the_register_window_reaches_the_evaluation_block() -> None:
    from pccm_builder.spec_loader import SpecError

    def edit(block):
        block["row_window"] = int(block["evaluation"]["heading_row"])
    with pytest.raises(SpecError, match="evaluation"):
        _mutate_manifest(edit)


def test_62_mutation_info_becomes_actionable() -> None:
    from pccm_builder.spec_loader import SpecError

    def edit(block):
        block["actionable_severities"] = list(block["severity_order"])
    with pytest.raises(SpecError, match="non-actionable"):
        _mutate_manifest(edit)


def test_63_mutation_the_structural_group_stops_sorting_first() -> None:
    from pccm_builder.spec_loader import SpecError

    def edit(block):
        block["group_order"] = ["Inputs", "Structure", "Calculation", "Simulation",
                                "Annual", "Sensitivity"]
    with pytest.raises(SpecError, match="first in both declared orders"):
        _mutate_manifest(edit)


# ===========================================================================
# H. THE SHEET SAYS ONLY WHAT IT DOES
# ===========================================================================
def test_64_the_placeholder_text_is_gone() -> None:
    from openpyxl import load_workbook

    stale = ("Not implemented yet", "no rules are implemented in Phase 1",
             "Reserved for the aggregate", "Reserved for the check register")
    # SCOPED TO THIS SHEET. `_SimData` is still legitimately unimplemented and
    # still says so; a manifest-wide ban would have convicted an honest note.
    plan = _plan()
    entry = next(sheet for sheet in _manifest()["sheets"] if sheet["name"] == plan.sheet)
    manifest_text = yaml.safe_dump(
        {"sheet": entry, "shell": _manifest()["phase9_shell"]}, sort_keys=False)
    workbook = load_workbook(_built_tree() / "PCCM_stageA.xlsx")
    try:
        strings = [cell.value for row in workbook[plan.sheet].iter_rows()
                   for cell in row
                   if isinstance(cell.value, str) and not cell.value.startswith("=")]
    finally:
        workbook.close()
    joined = "\n".join(strings)
    for phrase in stale:
        assert phrase not in joined, f"the built sheet still says: {phrase!r}"
        assert phrase not in manifest_text, f"the manifest still says: {phrase!r}"
    assert "Overall Model Status" in strings and "Validation Results" in strings


def test_65_the_sheet_claims_nothing_a_later_phase_owns() -> None:
    from openpyxl import load_workbook

    workbook = load_workbook(_built_tree() / "PCCM_stageA.xlsx")
    try:
        lowered = "\n".join(
            cell.value for row in workbook[_plan().sheet].iter_rows() for cell in row
            if isinstance(cell.value, str) and not cell.value.startswith("=")).lower()
    finally:
        workbook.close()
    for claim in ("sign-off", "hardening", "remediation plan", "chart", "tornado",
                  "histogram", "s-curve"):
        assert claim not in lowered, f"Model Check implies {claim!r} exists"


def test_66_no_new_command_button_was_introduced() -> None:
    from pccm_builder import load_structure_contract

    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    sheets = {button.sheet for button in structure.buttons}
    assert _plan().sheet not in sheets, "Phase 9 put a command button on Model Check"


def test_67_the_evaluation_source_is_on_the_sheet_and_not_hidden() -> None:
    """A summary whose working the reader cannot see is asking to be trusted
    rather than checked."""
    from openpyxl import load_workbook

    plan = _plan()
    workbook = load_workbook(_built_tree() / "PCCM_stageA.xlsx")
    try:
        sheet = workbook[plan.sheet]
        assert sheet.sheet_state == "visible"
        hidden = [letter for letter, dimension in sheet.column_dimensions.items()
                  if dimension.hidden]
        assert not hidden, f"the evaluation source is hidden behind {hidden}"
        assert sheet[f"{plan.label_column}{plan.evaluation['heading_row']}"].value == \
            plan.evaluation["heading"]
    finally:
        workbook.close()


# THE PRESENTATION KEYS THE P10-UX BATCH IS ALLOWED TO HAVE MOVED, and nothing
# else. Named per level so a range, a source, a state word or a row cannot hide
# among them.
UX_CHART_KEYS = {"anchor", "width", "height",
                 "value_axis_format", "category_axis_format"}


def test_68_the_accepted_phase_6_to_8_geometry_did_not_move() -> None:
    """PHASE 9 MAY ADD A BLOCK; it may not move a byte of a surface an accepted
    Windows run was produced against.

    RESTATED AT P10-UX, AND THE CLAIM IS THE SAME ONE. A human opened the real
    workbook for the first time and found the Dashboard charts too small to read
    and rows of #N/A down their axes. Correcting that moves chart GEOMETRY and
    AXIS PRESENTATION, which is exactly what this control was never about: it
    exists to stop a source range, a series, a row, a state word or a bridge
    address moving under an accepted Windows run. So the comparison is now
    field-for-field with the presentation keys NAMED, and every other byte of
    the accepted shell - including every range in it - is still required to be
    identical.
    """
    accepted = subprocess.run(["git", "show", "32e441e:pccm/spec/workbook.yaml"],
                              cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE,
                              text=True).stdout
    before = yaml.safe_load(accepted)
    now = _manifest()
    was, current = before["phase6_shell"], now["phase6_shell"]
    assert set(was) == set(current), sorted(set(was) ^ set(current))
    for key in was:
        if key != "charts":
            assert was[key] == current[key], f"the accepted {key} block moved"

    was_charts, now_charts = dict(was["charts"]), dict(current["charts"])
    # THE NEW TOP-LEVEL KEY IS NAMED AND IS PRESENTATION ONLY.
    assert set(now_charts) - set(was_charts) == {"axis_presentation"}
    assert set(now_charts["axis_presentation"]) == {
        "label_font_size", "category_label_position", "display_blanks_as"}
    # ONE NEW NUMBER FORMAT, AND EVERY ACCEPTED ONE UNCHANGED.
    assert set(now_charts["number_formats"]) - set(was_charts["number_formats"]) == {
        "money_axis"}
    for name, code in was_charts["number_formats"].items():
        assert now_charts["number_formats"][name] == code, name
    # EVERY CHART: same key, kind, source, categories, series and state
    # qualifiers. Only the named presentation keys may differ.
    was_by_key = {c["key"]: c for c in was_charts["charts"]}
    now_by_key = {c["key"]: c for c in now_charts["charts"]}
    assert set(was_by_key) == set(now_by_key)
    for key, chart in was_by_key.items():
        after = now_by_key[key]
        assert set(after) - set(chart) <= UX_CHART_KEYS, key
        for field in set(chart) - UX_CHART_KEYS:
            assert chart[field] == after[field], (key, field)
    # AND THE BRIDGE: same rows, same columns, same letters, same formats. The
    # only new key is `absent`, and only on a column a chart uses as its
    # CATEGORY - which is a label, not a point.
    categories = {c["categories"] for c in was_charts["charts"]}
    for name in ("annual", "distribution", "drivers"):
        was_block, now_block = was_charts["bridge"][name], now_charts["bridge"][name]
        assert set(was_block) == set(now_block), name
        for field in set(was_block) - {"columns"}:
            assert was_block[field] == now_block[field], (name, field)
        for before_col, after_col in zip(was_block["columns"], now_block["columns"]):
            extra = set(after_col) - set(before_col)
            assert extra <= {"absent"}, (name, extra)
            if extra:
                assert after_col["key"] in categories, (name, after_col["key"])
                assert after_col["absent"] == "blank", after_col
            for field in before_col:
                assert before_col[field] == after_col[field], (name, field)
    assert "phase9_shell" not in before
    assert before["workbook"]["locked_sheet_order"] == now["workbook"]["locked_sheet_order"]


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
