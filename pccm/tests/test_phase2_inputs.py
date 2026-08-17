#!/usr/bin/env python3
"""PCCM Phase 2 tests: functional Setup and Configuration layer.

Runs standalone or under pytest.

The locked input keys, defaults, table names and constants are re-declared here
independently of the input contract. The contract is the builder's authority;
this file is the conformance check, so contract drift fails here even when the
build itself succeeds.
"""

from __future__ import annotations

import os
import sys
import tempfile
import zipfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from openpyxl import load_workbook  # noqa: E402

from pccm_builder import build_workbook, load_contract, load_spec  # noqa: E402

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"
CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"

# --- locked Phase 2 decisions ----------------------------------------------
SETUP_INPUTS = {
    "project_name": ("Project Name", "inpProjectName", None),
    "duration_years": ("Project Duration (Years)", "inpDurationYears", None),
    "project_start_year": ("Project Start Year", "inpProjectStartYear", None),
    "base_year": ("Base Year", "inpBaseYear", None),
    "reporting_currency": ("Reporting Currency", "inpReportingCurrency", "SAR"),
    "selected_confidence_level": ("Selected Confidence Level", "inpSelectedConfidenceLevel", "P50"),
    "monte_carlo_iterations": ("Monte Carlo Iterations", "inpMonteCarloIterations", 10000),
    "discount_rate": ("Discount Rate", "inpDiscountRate", None),
    "random_seed": ("Random Seed", "inpRandomSeed", None),
}

CONFIG_TABLES = {
    "tblCategories": "lstCategories",
    "tblCurrencies": "lstCurrencies",
    "tblUOM": "lstUOM",
    "tblInflationProfiles": "lstInflationProfiles",
    "tblDistributions": "lstDistributions",
    "tblConfidenceLevels": "lstConfidenceLevels",
}

FX_TABLE = "tblFXRates"
FX_COLUMNS = ["Currency", "FX to SAR"]
LOCKED_DISTRIBUTIONS = ["Triangular", "Beta-PERT", "Uniform"]
LOCKED_CONFIDENCE_LEVELS = [f"P{n}" for n in range(50, 100, 5)]
MINIMUM_ITERATIONS = "1000"

_CACHE: dict[str, Path] = {}
_TEMPDIR: tempfile.TemporaryDirectory | None = None


# ---------------------------------------------------------------------------
def _artifact() -> Path:
    global _TEMPDIR
    if "primary" in _CACHE:
        return _CACHE["primary"]
    if _TEMPDIR is None:
        _TEMPDIR = tempfile.TemporaryDirectory(prefix="pccm-phase2-")
    previous = os.environ.get("PCCM_BUILD_TIMESTAMP")
    os.environ["PCCM_BUILD_TIMESTAMP"] = "1970-01-01 00:00:00 UTC"
    try:
        workbook, _ = build_workbook(load_spec(SPEC_PATH), load_contract(CONTRACT_PATH))
        path = Path(_TEMPDIR.name) / "stage_a.xlsx"
        workbook.save(path)
        workbook.close()
    finally:
        if previous is None:
            os.environ.pop("PCCM_BUILD_TIMESTAMP", None)
        else:
            os.environ["PCCM_BUILD_TIMESTAMP"] = previous
    _CACHE["primary"] = path
    return path


def _wb():
    return load_workbook(_artifact())


def _validations(worksheet) -> list:
    return list(worksheet.data_validations.dataValidation)


def _dv_for(worksheet, address: str):
    for dv in _validations(worksheet):
        if address in str(dv.sqref):
            return dv
    return None


# --- 1-2. Setup inputs and defined names -----------------------------------
def test_01_all_nine_setup_inputs_exist() -> None:
    contract = load_contract(CONTRACT_PATH)
    workbook = _wb()
    setup = workbook["Setup"]
    assert len(contract.inputs) == 9, f"expected 9 Setup inputs, found {len(contract.inputs)}"
    for key, (label, _, _) in SETUP_INPUTS.items():
        assert key in contract.inputs, f"contract is missing input {key!r}"
        spec = contract.inputs[key]
        assert spec.label == label, f"{key}: label is {spec.label!r}, expected {label!r}"
        assert setup[spec.label_cell].value == label, f"{key}: label not written to {spec.label_cell}"


def test_02_setup_defined_names_resolve_to_intended_cells() -> None:
    contract = load_contract(CONTRACT_PATH)
    workbook = _wb()
    for key, (_, defined_name, _) in SETUP_INPUTS.items():
        spec = contract.inputs[key]
        assert spec.defined_name == defined_name, (
            f"{key}: contract name {spec.defined_name!r}, lock says {defined_name!r}"
        )
        assert defined_name in workbook.defined_names, f"missing defined name {defined_name}"
        expected = f"'{spec.sheet}'!${spec.cell[0]}${spec.cell[1:]}"
        assert workbook.defined_names[defined_name].attr_text == expected, (
            f"{defined_name} -> {workbook.defined_names[defined_name].attr_text}, expected {expected}"
        )


# --- 3-7. locked defaults ---------------------------------------------------
def _default_of(key: str):
    contract = load_contract(CONTRACT_PATH)
    return _wb()["Setup"][contract.inputs[key].cell].value


def test_03_reporting_currency_is_sar_and_model_controlled() -> None:
    contract = load_contract(CONTRACT_PATH)
    spec = contract.inputs["reporting_currency"]
    assert spec.default == "SAR"
    assert spec.editable is False, "Reporting Currency must be model-controlled"
    assert spec.validation is None, "Reporting Currency must not have a dropdown"
    assert _default_of("reporting_currency") == "SAR"


def test_04_selected_confidence_level_defaults_to_p50() -> None:
    assert _default_of("selected_confidence_level") == "P50"


def test_05_monte_carlo_iterations_default_is_10000() -> None:
    assert _default_of("monte_carlo_iterations") == 10000


def test_06_random_seed_is_blank() -> None:
    assert _default_of("random_seed") is None


def test_07_discount_rate_is_blank() -> None:
    assert _default_of("discount_rate") is None


# --- 8-12. Config tables and constants --------------------------------------
def test_08_all_six_config_tables_exist() -> None:
    tables = getattr(_wb()["Config"], "tables", {})
    for name in CONFIG_TABLES:
        assert name in tables, f"Config is missing Excel Table {name}"
    assert set(tables) == set(CONFIG_TABLES), f"unexpected Config tables: {set(tables) - set(CONFIG_TABLES)}"


def test_09_currency_table_contains_sar() -> None:
    contract = load_contract(CONTRACT_PATH)
    table = next(t for t in contract.all_tables if t.table_name == "tblCurrencies")
    values = _column_values(_wb()["Config"], table)
    assert "SAR" in values, f"tblCurrencies does not contain SAR: {values}"


def test_10_currency_table_has_no_fx_rate_field() -> None:
    contract = load_contract(CONTRACT_PATH)
    table = next(t for t in contract.all_tables if t.table_name == "tblCurrencies")
    headers = [c.header for c in table.columns]
    assert headers == ["Currency"], f"tblCurrencies schema is {headers}, expected ['Currency']"


def test_11_distribution_constants_are_exact() -> None:
    contract = load_contract(CONTRACT_PATH)
    table = next(t for t in contract.all_tables if t.table_name == "tblDistributions")
    assert _column_values(_wb()["Config"], table) == LOCKED_DISTRIBUTIONS
    assert table.editable is False, "distribution list must be a locked constant"


def test_12_confidence_level_constants_are_exact() -> None:
    contract = load_contract(CONTRACT_PATH)
    table = next(t for t in contract.all_tables if t.table_name == "tblConfidenceLevels")
    assert _column_values(_wb()["Config"], table) == LOCKED_CONFIDENCE_LEVELS
    assert table.editable is False, "confidence level list must be a locked constant"


def _column_values(worksheet, table, index: int = 0) -> list:
    letter = table.column_letter(index)
    values = [
        worksheet[f"{letter}{row}"].value
        for row in range(table.first_data_row, table.last_data_row + 1)
    ]
    return [v for v in values if v is not None]


# --- 13-14. FX table --------------------------------------------------------
def test_13_setup_fx_table_exists_with_exact_schema() -> None:
    contract = load_contract(CONTRACT_PATH)
    assert FX_TABLE in getattr(_wb()["Setup"], "tables", {}), "Setup is missing tblFXRates"
    table = next(t for t in contract.all_tables if t.table_name == FX_TABLE)
    assert [c.header for c in table.columns] == FX_COLUMNS


def test_14_sar_fx_identity_row_is_one() -> None:
    contract = load_contract(CONTRACT_PATH)
    table = next(t for t in contract.all_tables if t.table_name == FX_TABLE)
    setup = _wb()["Setup"]
    currency = setup[f"{table.column_letter(0)}{table.first_data_row}"].value
    rate = setup[f"{table.column_letter(1)}{table.first_data_row}"].value
    assert currency == "SAR", f"first FX row is {currency!r}, expected SAR"
    assert rate == 1, f"SAR identity rate is {rate!r}, expected 1"


# --- 15. list defined names -------------------------------------------------
def test_15_required_list_defined_names_exist() -> None:
    workbook = _wb()
    contract = load_contract(CONTRACT_PATH)
    for table_name, list_name in CONFIG_TABLES.items():
        assert list_name in workbook.defined_names, f"missing defined name {list_name}"
        table = next(t for t in contract.all_tables if t.table_name == table_name)
        assert workbook.defined_names[list_name].attr_text == table.absolute_data_range(), (
            f"{list_name} does not cover the {table_name} data body"
        )


# --- 16-19. data validation -------------------------------------------------
def test_16_confidence_level_validation_uses_the_list() -> None:
    contract = load_contract(CONTRACT_PATH)
    setup = _wb()["Setup"]
    dv = _dv_for(setup, contract.inputs["selected_confidence_level"].cell)
    assert dv is not None, "Selected Confidence Level has no data validation"
    assert dv.type == "list"
    assert dv.formula1 == "=lstConfidenceLevels", f"formula1 is {dv.formula1!r}"


def test_17_iterations_validation_enforces_minimum_1000() -> None:
    contract = load_contract(CONTRACT_PATH)
    setup = _wb()["Setup"]
    dv = _dv_for(setup, contract.inputs["monte_carlo_iterations"].cell)
    assert dv is not None, "Monte Carlo Iterations has no data validation"
    assert dv.type == "whole"
    assert dv.operator == "greaterThanOrEqual"
    assert dv.formula1 == MINIMUM_ITERATIONS, f"formula1 is {dv.formula1!r}"


def test_18_fx_currency_validation_uses_the_currency_master() -> None:
    contract = load_contract(CONTRACT_PATH)
    table = next(t for t in contract.all_tables if t.table_name == FX_TABLE)
    dv = _dv_for(_wb()["Setup"], table.data_range(0))
    assert dv is not None, "FX Currency column has no data validation"
    assert dv.type == "list"
    assert dv.formula1 == "=lstCurrencies", f"formula1 is {dv.formula1!r}"


def test_19_fx_rate_validation_requires_positive_when_populated() -> None:
    contract = load_contract(CONTRACT_PATH)
    table = next(t for t in contract.all_tables if t.table_name == FX_TABLE)
    dv = _dv_for(_wb()["Setup"], table.data_range(1))
    assert dv is not None, "FX rate column has no data validation"
    assert dv.type == "decimal"
    assert dv.operator == "greaterThan"
    assert dv.formula1 == "0"
    assert dv.allow_blank, "blank unused FX rows must remain acceptable"


def test_19b_duration_validation_is_positive_and_uncapped() -> None:
    contract = load_contract(CONTRACT_PATH)
    dv = _dv_for(_wb()["Setup"], contract.inputs["duration_years"].cell)
    assert dv is not None, "Project Duration has no data validation"
    assert dv.type == "whole"
    assert dv.operator == "greaterThanOrEqual"
    assert dv.formula1 == "1"
    assert dv.formula2 in (None, ""), "Project Duration must not be capped"


def test_19c_no_unsupported_validation_invented() -> None:
    """Start Year, Base Year, Discount Rate and Random Seed have no business limits yet."""
    contract = load_contract(CONTRACT_PATH)
    setup = _wb()["Setup"]
    for key in ("project_start_year", "base_year", "discount_rate", "random_seed", "project_name"):
        assert contract.inputs[key].validation is None, f"{key} declares a validation rule"
        assert _dv_for(setup, contract.inputs[key].cell) is None, f"{key} has an applied validation"


# --- 20-23. nothing from a later phase --------------------------------------
def test_20_no_cost_line_or_risk_tables_yet() -> None:
    workbook = _wb()
    for sheet in ("Cost Lines", "Risk Register", "Inflation", "Cost Profiling", "Risk Profiling"):
        assert not getattr(workbook[sheet], "tables", {}), f"{sheet} already declares a table"
    all_tables = {n for ws in workbook.worksheets for n in getattr(ws, "tables", {})}
    assert all_tables == set(CONFIG_TABLES) | {FX_TABLE}, f"unexpected tables: {all_tables}"


def test_21_no_formulas_or_business_calculations() -> None:
    workbook = _wb()
    offenders = [
        f"{ws.title}!{cell.coordinate}"
        for ws in workbook.worksheets
        for row in ws.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    ]
    assert not offenders, f"formulas present: {offenders[:10]}"


def test_22_no_vba_project() -> None:
    with zipfile.ZipFile(_artifact()) as archive:
        names = [n.lower() for n in archive.namelist()]
    assert not any(n.endswith("vbaproject.bin") for n in names)
    assert _artifact().suffix == ".xlsx"


def test_23_no_external_links_or_connections() -> None:
    with zipfile.ZipFile(_artifact()) as archive:
        names = archive.namelist()
    assert not [n for n in names if "externalLink" in n]
    assert not [n for n in names if "connections" in n.lower()]
    assert not [n for n in names if "queryTable" in n]


# --- 24-25. Phase 1 guarantees still hold ------------------------------------
def test_24_phase1_sheet_order_and_visibility_still_hold() -> None:
    workbook = _wb()
    spec = load_spec(SPEC_PATH)
    assert workbook.sheetnames == spec.workbook["locked_sheet_order"]
    assert workbook.active.title == "Dashboard"
    assert workbook["_Calc"].sheet_state == "hidden"
    assert workbook["_SimData"].sheet_state == "veryHidden"
    for name in workbook.sheetnames:
        if name not in ("_Calc", "_SimData"):
            assert workbook[name].sheet_state == "visible", f"{name} is not visible"


def test_25_artifact_reopens() -> None:
    workbook = load_workbook(_artifact())
    assert len(workbook.sheetnames) == 14
    workbook.close()


# --- 26-28. contract integrity ----------------------------------------------
def test_26_contract_defined_names_are_unique() -> None:
    contract = load_contract(CONTRACT_PATH)
    names = list(contract.input_defined_names) + list(contract.list_defined_names)
    assert len(names) == len(set(names)), "duplicate defined names in the contract"
    tables = [t.table_name for t in contract.all_tables]
    assert len(tables) == len(set(tables)), "duplicate table names in the contract"
    assert not set(names) & set(tables), "a name is used for both a range and a table"


def test_27_contract_targets_are_valid() -> None:
    contract = load_contract(CONTRACT_PATH)
    spec = load_spec(SPEC_PATH)
    known = set(spec.sheet_names)
    workbook = _wb()
    for input_spec in contract.inputs.values():
        assert input_spec.sheet in known, f"{input_spec.key} targets unknown sheet"
        assert workbook[input_spec.sheet][input_spec.cell] is not None
    for table in contract.all_tables:
        assert table.sheet in known, f"{table.table_name} targets unknown sheet"
        assert table.table_name in getattr(workbook[table.sheet], "tables", {})


def test_28_setup_is_the_only_fx_rate_owner() -> None:
    contract = load_contract(CONTRACT_PATH)
    rate_tables = [
        (t.sheet, t.table_name)
        for t in contract.all_tables
        if any(
            "fx" in c.header.lower() or "rate" in c.header.lower() for c in t.columns
        )
    ]
    assert rate_tables == [("Setup", FX_TABLE)], f"FX-rate owners: {rate_tables}"
    config_headers = [
        c.header
        for t in contract.all_tables
        if t.sheet == "Config"
        for c in t.columns
    ]
    assert not [h for h in config_headers if "fx" in h.lower() or "rate" in h.lower()], (
        f"Config declares a rate-like column: {config_headers}"
    )


def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 2 tests - functional Setup & Configuration layer")
    print("=" * 70)
    for name, fn in tests:
        try:
            fn()
        except AssertionError as error:
            failures += 1
            print(f"  [FAIL] {name}\n         {error}")
        except Exception as error:  # noqa: BLE001
            failures += 1
            print(f"  [ERROR] {name}\n          {type(error).__name__}: {error}")
        else:
            print(f"  [PASS] {name}")
    print("=" * 70)
    print(f"  {len(tests) - failures} passed, {failures} failed")
    if _TEMPDIR is not None:
        _TEMPDIR.cleanup()
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(_run_all())
