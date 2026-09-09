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


def test_46c_the_permanent_id_is_only_in_prose_and_that_is_reported() -> None:
    """THE DECLARED GAP, STATED AS A FACT ABOUT THE SOURCE.

    The owner names the driver when it refuses - and it builds that name from
    ResolvedDriver.PermanentId, which it holds. What it publishes is a STRING.
    Three modules write such sentences in four different shapes, so no parser
    could be written that a fifth shape would not silently break. This control
    exists so the day a structured field appears, it fails and is removed.
    """
    check = _code("modCalcCheck")
    assert "Private Function DriverLabel(ByRef driver As ResolvedDriver) As String" in check
    assert "driver.PermanentId" in check, "the owner does hold the id"
    # THE PUBLISHED SHAPE IS A SENTENCE AND NOTHING ELSE.
    assert ("Public Function CheckResolvedModel(ByRef model As ResolvedModel, _\n"
            "                                   ByRef detail As String) As Boolean") in check
    assert "ByRef subject" not in check and "ByRef offending" not in check.lower()
    # AND NO OWNER ANYWHERE PUBLISHES THE ID ON ITS OWN.
    for module in sorted(SRC.glob("*.bas")):
        code = _code(module.stem)
        for structured in ("OffendingId", "RefusalSubject", "FailingDriverId"):
            assert structured not in code, (
                f"{module.name} publishes {structured}; the gap is closed and this "
                "control should be replaced by the real assertion")
    # AND MODEL CHECK PARSES NOTHING. No formula on the sheet takes the sentence
    # apart - the only text-splitting on the surface is the structural report's
    # own line separator, which the owner writes deliberately.
    plan = _plan()
    detail = plan.reading_cell("calculation_refusal_detail")
    for value in _cells(plan).values():
        if isinstance(value, str) and detail in value:
            for splitter in ("FIND(", "MID(", "LEFT(", "RIGHT(", "SEARCH("):
                assert splitter not in value, (
                    f"a Model Check formula takes the refusal sentence apart: {value}")


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
    assert "modCalcReport.CalcReportDerivedStatus(detail)" in adapter
    assert "CVErr(xlErrValue)" in adapter, "the adapter must fail loud, not wrong"
    assert "PCCM_CalculationStatus" not in adapter
    assert "PCCM_CalculationAttemptDetail" not in adapter, (
        "the live reason may not be taken from the persisted attempt")
    # AND THE OWNER HANDS BACK WHAT IT ALREADY WROTE; nothing is re-derived.
    exposure = _procedure("modCalcReport", "CalcReportDerivedStatus")
    assert "ByRef detail As String" in exposure
    assert "Optional" not in exposure, (
        "a typed Optional with no default is a VBA compile error, and both "
        "callers supply the argument")
    assert "PrepareCurrentCalculation(package, detail)" in exposure
    assert "WriteStatusBlock" not in exposure


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
    assert "Public Function CalcReportDerivedStatus(ByRef detail As String) As String" in calc
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
    assert "modCalcReport.CalcReportDerivedStatus(detail)" in adapter
    assert "CVErr(xlErrValue)" in adapter, "the adapter must fail loud, not wrong"
    assert "PCCM_CalculationStatus" not in adapter


def test_45_the_adapter_invents_no_state_word() -> None:
    """Phase 9 owns no vocabulary. The four calculation words arrive spelled the
    way modCalcReport spells them, and modResultsState defines none of them."""
    adapter = _procedure("modResultsState", "PCCM_ModelCheckCalculationState")
    for word in ("NOT CALCULATED", "CURRENT", "STALE", "INVALID", "PASS", "WARNING",
                 "ERROR"):
        assert f'"{word}"' not in adapter, f"the adapter spells {word!r}"


def test_46_the_new_source_is_purely_additive() -> None:
    """A declaration buys the right to ADD to an accepted module, never to
    rewrite one. Checked against the Phase-7 acceptance head."""
    head = "ad78988"
    for module in ("modCalcReport.bas", "modResultsState.bas"):
        path = f"pccm/src/vba/{module}"
        # AGAINST THE WORKING TREE, not against HEAD: a control that only ever
        # sees committed history cannot refuse the edit that is being made.
        diff = subprocess.run(["git", "diff", head, "--", path],
                              cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE,
                              text=True).stdout
        removed = [line for line in diff.splitlines()
                   if line.startswith("-") and not line.startswith("---")]
        assert not removed, f"{module} removes {len(removed)} line(s): {removed[:3]}"


def test_47_no_module_silently_changed_its_line_endings() -> None:
    """A DEFECT THIS STEP ACTUALLY MADE, AND THE CONTROL THAT NOW CATCHES IT.

    `modCalcReport.bas` is CRLF throughout and an edit written with LF newlines
    rewrote every line of it. Nothing about the source read differently; the diff
    was the whole file, the additive guarantee was gone, and VBA injection expects
    the separators the module was written with. So the convention itself is now
    part of what may not change.
    """
    head = "ad78988"
    for module in sorted(SRC.glob("*.bas")):
        path = f"pccm/src/vba/{module.name}"
        before = subprocess.run(["git", "show", f"{head}:{path}"], cwd=REPO_ROOT,
                                capture_output=True)
        if before.returncode != 0:
            continue          # added after the acceptance head; it sets its own
        raw = module.read_bytes()
        was_crlf = b"\r\n" in before.stdout
        is_crlf = b"\r\n" in raw
        assert was_crlf == is_crlf, (
            f"{module.name} changed line-ending convention: "
            f"{'CRLF' if was_crlf else 'LF'} -> {'CRLF' if is_crlf else 'LF'}")
        # AND IT IS NOT MIXED, which is what a careless append leaves behind.
        assert raw.count(b"\n") == (raw.count(b"\r\n") if is_crlf else raw.count(b"\n")), \
            f"{module.name} mixes line endings"
        if is_crlf:
            assert raw.count(b"\n") == raw.count(b"\r\n"), (
                f"{module.name} mixes LF lines into a CRLF module")


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


def test_68_the_accepted_phase_6_to_8_geometry_did_not_move() -> None:
    """PHASE 9 MAY ADD A BLOCK; it may not move a byte of a surface an accepted
    Windows run was produced against."""
    accepted = subprocess.run(["git", "show", "32e441e:pccm/spec/workbook.yaml"],
                              cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE,
                              text=True).stdout
    before = yaml.safe_load(accepted)
    now = _manifest()
    assert before["phase6_shell"] == now["phase6_shell"], (
        "the accepted Phase-6/7/8 shell changed in this step")
    assert "phase9_shell" not in before
    assert before["workbook"]["locked_sheet_order"] == now["workbook"]["locked_sheet_order"]


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
