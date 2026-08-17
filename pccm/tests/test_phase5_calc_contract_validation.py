#!/usr/bin/env python3
"""PCCM Phase 5 Gate-A Step-1 negative tests for the calculation contract.

`spec/calc_contract.yaml` is the narrowly scoped fifth authority. Two things have
to be enforced, not merely documented:

  LAYOUT   the contract must ENCODE the accepted Revision-E `_Calc` layout, not
           choose a new one. Every anchor, band width, row range and schema is
           checked against a locked constant in `calc_loader.py`.

  BOUNDARY the contract must not acquire authority it was not given. The
           fingerprint hash mathematics in particular is owned by exactly one
           source - `builder/pccm_builder/calc_fingerprint.py` - and a second
           hand-maintained copy in YAML is the drift this architecture exists to
           prevent.

A malformed contract, or one that disagrees with the four accepted
specifications, must fail loudly rather than produce a `_Calc` sheet whose
calculation record is quietly in the wrong place.

NO VBA IS EXECUTED HERE. This is Linux static validation (plan section 21.0).

Runs standalone or under pytest.
"""

from __future__ import annotations

import copy
import dataclasses
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable

import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    CalcContractError,
    load_calc_contract,
    load_contract,
    load_driver_contract,
    load_spec,
    load_structure_contract,
)
from pccm_builder.calc_loader import (  # noqa: E402
    LOCKED_ANNUAL_HEADERS,
    LOCKED_ATTEMPT_RESULT,
    LOCKED_CALC_STATE_ROWS,
    LOCKED_CALC_TOTALS_ROWS,
    LOCKED_DERIVED_STATUS,
    LOCKED_DRIVER_HEADERS,
    LOCKED_FP_VERSION,
    LOCKED_PHASE4_CELLS,
    LOCKED_TABLE_ANCHORS,
    validate_calc_against,
)

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"
CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
DRIVERS_PATH = PCCM_ROOT / "spec" / "driver_contract.yaml"
STRUCTURE_PATH = PCCM_ROOT / "spec" / "structure_contract.yaml"
CALC_PATH = PCCM_ROOT / "spec" / "calc_contract.yaml"


def _base() -> dict[str, Any]:
    with CALC_PATH.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def _write(data: dict[str, Any], tmp: str, name: str = "broken.yaml") -> Path:
    path = Path(tmp) / name
    with path.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(data, handle, sort_keys=False, allow_unicode=True)
    return path


def _rejected(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    """The mutated contract must fail at load time."""
    data = copy.deepcopy(_base())
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-badcalc-") as tmp:
        path = _write(data, tmp)
        try:
            load_calc_contract(path)
        except CalcContractError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of CalcContractError"
            ) from error
    raise AssertionError(f"{reason}: an invalid calculation contract was silently accepted")


def _rejected_cross(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    """The mutated contract must fail cross-contract validation."""
    data = copy.deepcopy(_base())
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-badcalc-") as tmp:
        calc = load_calc_contract(_write(data, tmp))
        try:
            validate_calc_against(
                calc,
                load_spec(SPEC_PATH),
                load_contract(CONTRACT_PATH),
                load_driver_contract(DRIVERS_PATH),
                load_structure_contract(STRUCTURE_PATH),
            )
        except CalcContractError:
            return
    raise AssertionError(f"{reason}: silently accepted")


def _rejected_cross_replace(*, reason: str, **fields: Any) -> None:
    """Cross-validation must reject a contract altered AFTER loading.

    Several cross-checks are unreachable through the YAML because the load-time
    layout lock refuses the edit first. That ordering is correct - the earliest
    guard should win - but it would leave the later guard untested, so these cases
    replace the field on an already-loaded contract instead.
    """
    calc = dataclasses.replace(load_calc_contract(CALC_PATH), **fields)
    try:
        validate_calc_against(
            calc,
            load_spec(SPEC_PATH),
            load_contract(CONTRACT_PATH),
            load_driver_contract(DRIVERS_PATH),
            load_structure_contract(STRUCTURE_PATH),
        )
    except CalcContractError:
        return
    raise AssertionError(f"{reason}: silently accepted")


def _state_field(data: dict[str, Any], key: str) -> dict[str, Any]:
    return next(f for f in data["scalar_blocks"]["calc_state"]["fields"] if f["key"] == key)


def _table(data: dict[str, Any], table_name: str) -> dict[str, Any]:
    return next(t for t in data["tables"].values() if t["table_name"] == table_name)


# ---------------------------------------------------------------------------
# the real contract
# ---------------------------------------------------------------------------
def test_the_real_calc_contract_loads_and_agrees_with_the_other_four() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert calc.sheet == "_Calc"
    assert calc.fingerprint_version == LOCKED_FP_VERSION
    assert len(calc.all_tables) == 5
    assert len(calc.calc_state.fields) == 8
    assert len(calc.calc_totals.fields) == 10
    validate_calc_against(
        calc,
        load_spec(SPEC_PATH),
        load_contract(CONTRACT_PATH),
        load_driver_contract(DRIVERS_PATH),
        load_structure_contract(STRUCTURE_PATH),
    )


def test_a_missing_calc_contract_is_reported_not_defaulted() -> None:
    try:
        load_calc_contract(PCCM_ROOT / "spec" / "does_not_exist.yaml")
    except CalcContractError:
        return
    raise AssertionError("a missing calculation contract was silently accepted")


# ---------------------------------------------------------------------------
# the accepted anchors
# ---------------------------------------------------------------------------
def test_every_accepted_anchor_is_encoded_exactly() -> None:
    calc = load_calc_contract(CALC_PATH)
    for table_name, (first, last, header_row, width) in LOCKED_TABLE_ANCHORS.items():
        table = calc.table_by_name(table_name)
        assert table.first_column == first
        assert table.last_column == last
        assert table.header_row == header_row
        assert len(table.columns) == width
        assert table.band_width == width


def test_all_five_listobjects_anchor_at_the_same_header_row() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert {t.header_row for t in calc.all_tables} == {15}


def test_the_bands_are_separated_by_two_column_gutters() -> None:
    """The gutter is what turns a widened schema into a build failure.

    Without it, one extra column would silently overwrite the neighbouring table's
    first column instead of tripping the overlap assertion.
    """
    calc = load_calc_contract(CALC_PATH)
    ordered = sorted(calc.all_tables, key=lambda t: t.first_column_index)
    for earlier, later in zip(ordered, ordered[1:]):
        gap = later.first_column_index - earlier.last_column_index - 1
        assert gap == 2, f"{earlier.table_name} -> {later.table_name} gutter is {gap}"


def test_a_moved_band_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _table(data, "tblCalcFX")["first_column"] = "R"

    _rejected(mutate, "a table anchored away from its accepted column")


def test_a_moved_header_row_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _table(data, "tblCalcYears")["header_row"] = 16

    _rejected(mutate, "a table anchored at the wrong header row")


def test_a_schema_narrower_than_its_band_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _table(data, "tblCalcFX")["columns"].pop()

    _rejected(mutate, "a schema that does not fill its locked band")


def test_a_schema_wider_than_its_band_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        table = _table(data, "tblCalcFX")
        extra = copy.deepcopy(table["columns"][0])
        extra["key"] = "spare"
        extra["header"] = "Spare"
        table["columns"].append(extra)

    _rejected(mutate, "a schema wider than its locked band")


def test_overlapping_bands_are_rejected() -> None:
    """A band widened into its neighbour must fail, not silently overwrite it."""

    def mutate(data: dict[str, Any]) -> None:
        table = _table(data, "tblCalcFX")
        table["last_column"] = "X"
        for index in range(2):
            extra = copy.deepcopy(table["columns"][0])
            extra["key"] = f"spare_{index}"
            extra["header"] = f"Spare {index}"
            table["columns"].append(extra)

    _rejected(mutate, "two dynamic column bands overlapping")


def test_a_missing_listobject_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        del data["tables"]["calc_fx"]

    _rejected(mutate, "a required Phase-5 ListObject missing")


def test_an_unexpected_listobject_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        extra = copy.deepcopy(data["tables"]["calc_fx"])
        extra["table_name"] = "tblCalcExtra"
        data["tables"]["calc_extra"] = extra

    _rejected(mutate, "an undeclared Phase-5 ListObject")


def test_a_duplicate_column_header_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        columns = _table(data, "tblCalcYears")["columns"]
        columns[1]["header"] = columns[0]["header"]

    _rejected(mutate, "a duplicate column header")


# ---------------------------------------------------------------------------
# the Phase-4 reservation
# ---------------------------------------------------------------------------
def test_the_phase4_reservation_is_rows_one_to_eleven_and_the_counter_cells() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert (calc.phase4_first_row, calc.phase4_last_row) == (1, 11)
    assert calc.phase4_cells == LOCKED_PHASE4_CELLS


def test_no_phase5_block_touches_the_phase4_rows() -> None:
    calc = load_calc_contract(CALC_PATH)
    reserved = set(calc.phase4_reserved_rows)
    for block in calc.scalar_blocks.values():
        assert not set(block.rows) & reserved
    for table in calc.all_tables:
        assert table.header_row not in reserved


def test_a_scalar_block_intruding_on_phase4_rows_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        block = data["scalar_blocks"]["calc_state"]
        block["first_row"] = 8
        block["last_row"] = 15
        for offset, entry in enumerate(block["fields"]):
            entry["row"] = 8 + offset

    _rejected(mutate, "a Phase-5 block placed on frozen Phase-4 rows")


def test_a_narrowed_phase4_reservation_is_rejected() -> None:
    """Shrinking the reservation would 'free' rows Phase 4 still owns."""

    def mutate(data: dict[str, Any]) -> None:
        data["phase4_reservation"]["last_row"] = 9

    _rejected(mutate, "a narrowed Phase-4 reservation")


def test_dropping_a_reserved_counter_cell_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["phase4_reservation"]["cells"] = ["C10"]

    _rejected(mutate, "a Phase-4 counter cell dropped from the reservation")


def test_reserved_cells_that_are_not_the_counters_fail_cross_validation() -> None:
    """The second, independent guard.

    The load-time check already locks these cells against `calc_loader`'s own
    constant, so this cannot be reached by editing the YAML. It is reached by
    replacing the field on a loaded contract, which is the point: the cross-check
    asks a different question - are these the cells `structure_contract.yaml`
    actually declares? - and must hold on its own.
    """
    _rejected_cross_replace(
        phase4_cells=("C10", "C12"),
        reason="reserved cells that are not the Phase-4 counter cells",
    )


# ---------------------------------------------------------------------------
# calc_state
# ---------------------------------------------------------------------------
def test_calc_state_occupies_rows_thirteen_to_twenty() -> None:
    calc = load_calc_contract(CALC_PATH)
    block = calc.calc_state
    assert (block.first_row, block.last_row) == LOCKED_CALC_STATE_ROWS == (13, 20)
    assert [f.row for f in block.fields] == list(range(13, 21))


def test_calc_state_groups_make_the_commit_one_contiguous_range() -> None:
    calc = load_calc_contract(CALC_PATH)
    groups = [f.group for f in calc.calc_state.fields]
    assert groups == ["snapshot"] * 4 + ["attempt"] * 2 + ["derived"] * 2
    assert calc.calc_state.value_range() == "C13:C20"


def test_moving_calc_state_off_its_rows_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        block = data["scalar_blocks"]["calc_state"]
        block["first_row"] = 14
        block["last_row"] = 21
        for offset, entry in enumerate(block["fields"]):
            entry["row"] = 14 + offset

    _rejected(mutate, "calc_state moved off rows 13:20")


def test_reordering_calc_state_rows_is_rejected() -> None:
    """Row order is load-bearing, not cosmetic."""

    def mutate(data: dict[str, Any]) -> None:
        fields = data["scalar_blocks"]["calc_state"]["fields"]
        fields[3], fields[4] = fields[4], fields[3]
        for offset, entry in enumerate(fields):
            entry["row"] = 13 + offset

    _rejected(mutate, "a reordered calc_state block")


def test_a_seeded_fingerprint_version_is_rejected() -> None:
    """A never-calculated workbook must not look as though it holds a snapshot."""

    def mutate(data: dict[str, Any]) -> None:
        _state_field(data, "fingerprint_version")["initial"] = 1

    _rejected(mutate, "a build-time seeded Fingerprint Version")


def test_the_locked_initial_values_are_encoded() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert calc.initial_values() == {
        "C13": None,
        "C14": None,
        "C15": None,
        "C16": None,
        "C17": "NONE",
        "C18": None,
        "C19": "NOT CALCULATED",
        "C20": None,
    }


def test_seeding_a_blank_cell_with_a_value_is_rejected() -> None:
    """Blank means blank - never 0 and never an empty string."""

    def mutate(data: dict[str, Any]) -> None:
        _state_field(data, "last_successful_stamp")["initial"] = ""

    _rejected(mutate, "a seeded value in a cell that must be blank")


def test_a_calc_state_field_drawing_from_an_unknown_label_set_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _state_field(data, "last_attempt_result")["enum"] = "some_other_axis"

    _rejected(mutate, "a field bound to an unknown label set")


# ---------------------------------------------------------------------------
# calc_totals
# ---------------------------------------------------------------------------
def test_calc_totals_occupies_rows_twentythree_to_thirtytwo() -> None:
    calc = load_calc_contract(CALC_PATH)
    block = calc.calc_totals
    assert (block.first_row, block.last_row) == LOCKED_CALC_TOTALS_ROWS == (23, 32)
    assert [f.row for f in block.fields] == list(range(23, 33))
    assert block.value_range() == "C23:C32"


def test_calc_totals_pairs_each_measure_with_its_pv_row() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert [f.measure for f in calc.calc_totals.fields] == list("AABBCCDDEE")
    assert [f.basis for f in calc.calc_totals.fields] == ["nominal", "pv"] * 5


def test_moving_calc_totals_off_its_rows_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        block = data["scalar_blocks"]["calc_totals"]
        block["first_row"] = 24
        block["last_row"] = 33
        for offset, entry in enumerate(block["fields"]):
            entry["row"] = 24 + offset

    _rejected(mutate, "calc_totals moved off rows 23:32")


def test_a_seeded_total_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["scalar_blocks"]["calc_totals"]["fields"][0]["initial"] = 0

    _rejected(mutate, "a seeded zero total")


def test_a_total_in_a_currency_other_than_sar_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["scalar_blocks"]["calc_totals"]["fields"][0]["units"] = "USD"

    _rejected(mutate, "a headline total denominated outside the reporting currency")


def test_overlapping_scalar_blocks_are_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        block = data["scalar_blocks"]["calc_totals"]
        block["first_row"] = 13
        block["last_row"] = 22
        for offset, entry in enumerate(block["fields"]):
            entry["row"] = 13 + offset

    _rejected(mutate, "two scalar blocks claiming the same rows")


# ---------------------------------------------------------------------------
# the locked schemas
# ---------------------------------------------------------------------------
def test_the_driver_audit_schema_is_the_exact_twentyone_columns() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert calc.table_by_name("tblCalcDrivers").headers == list(LOCKED_DRIVER_HEADERS)
    assert len(LOCKED_DRIVER_HEADERS) == 21


def test_the_annual_schema_is_the_exact_eight_columns_including_calendar_year() -> None:
    calc = load_calc_contract(CALC_PATH)
    headers = calc.table_by_name("tblCalcAnnual").headers
    assert headers == list(LOCKED_ANNUAL_HEADERS)
    assert len(LOCKED_ANNUAL_HEADERS) == 8
    assert "Calendar Year" in headers


def test_a_renamed_driver_audit_column_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _table(data, "tblCalcDrivers")["columns"][11]["header"] = "K Nominal"

    _rejected(mutate, "a renamed driver audit column")


def test_a_reordered_driver_audit_schema_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        columns = _table(data, "tblCalcDrivers")["columns"]
        columns[7], columns[8] = columns[8], columns[7]

    _rejected(mutate, "a reordered driver audit schema")


def test_a_renamed_annual_column_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _table(data, "tblCalcAnnual")["columns"][1]["header"] = "Year"

    _rejected(mutate, "a renamed annual audit column")


def test_every_driver_audit_column_declares_the_kinds_it_applies_to() -> None:
    """"Blank, never zero" is only checkable if each column says who populates it."""
    calc = load_calc_contract(CALC_PATH)
    table = calc.table_by_name("tblCalcDrivers")
    for column in table.columns:
        assert column.applies_to, column.key
    quantity = next(c for c in table.columns if c.key == "quantity")
    probability = next(c for c in table.columns if c.key == "probability")
    assert quantity.applies_to == ("cost_line",)
    assert probability.applies_to == ("risk",)


def test_a_driver_audit_column_without_applies_to_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        del _table(data, "tblCalcDrivers")["columns"][7]["applies_to"]

    _rejected(mutate, "an audit column that does not say which kinds populate it")


# ---------------------------------------------------------------------------
# the two orthogonal state axes
# ---------------------------------------------------------------------------
def test_the_derived_status_axis_is_exactly_four_values() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert calc.derived_status_labels == LOCKED_DERIVED_STATUS
    assert calc.derived_status_labels == ("NOT CALCULATED", "CURRENT", "STALE", "INVALID")


def test_the_attempt_result_axis_is_exactly_four_values() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert calc.attempt_result_labels == LOCKED_ATTEMPT_RESULT
    assert calc.attempt_result_labels == ("NONE", "SUCCESS", "REFUSED", "FAILED")


def test_the_two_axes_share_no_label() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert not set(calc.derived_status_labels) & set(calc.attempt_result_labels)


def test_refused_on_the_derived_status_axis_is_rejected() -> None:
    """REFUSED is an ATTEMPT result. A model is INVALID whether or not anyone
    pressed Calculate, so the status axis must not carry it."""

    def mutate(data: dict[str, Any]) -> None:
        data["state_labels"]["derived_status"].append("REFUSED")

    _rejected(mutate, "REFUSED on the derived-status axis")


def test_failed_on_the_derived_status_axis_is_rejected() -> None:
    """FAILED is an ATTEMPT result. After a rolled-back write the model is STALE,
    derived from the inputs, not forced from the attempt."""

    def mutate(data: dict[str, Any]) -> None:
        data["state_labels"]["derived_status"].append("FAILED")

    _rejected(mutate, "FAILED on the derived-status axis")


def test_a_missing_derived_status_value_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["state_labels"]["derived_status"].remove("STALE")

    _rejected(mutate, "a derived-status value dropped")


def test_a_reordered_status_axis_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["state_labels"]["derived_status"] = ["CURRENT", "NOT CALCULATED", "STALE", "INVALID"]

    _rejected(mutate, "a reordered derived-status axis")


def test_a_missing_attempt_result_value_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["state_labels"]["attempt_result"].remove("NONE")

    _rejected(mutate, "an attempt-result value dropped")


# ---------------------------------------------------------------------------
# FP_VERSION and the authority boundary
# ---------------------------------------------------------------------------
def test_fp_version_is_one() -> None:
    assert load_calc_contract(CALC_PATH).fingerprint_version == 1
    assert LOCKED_FP_VERSION == 1


def test_a_changed_fp_version_is_rejected() -> None:
    """Bumping it declares a new canonical encoding and invalidates every stored
    digest. That is a design change, not a contract edit."""

    def mutate(data: dict[str, Any]) -> None:
        data["fingerprint"]["version"] = 2

    _rejected(mutate, "a silently bumped fingerprint version")


def test_the_real_contract_carries_no_hash_mathematics() -> None:
    """Scans the FILE TEXT, so a commented-out constant is caught too."""
    text = CALC_PATH.read_text(encoding="utf-8")
    import re

    for literal in ("131", "2147483647", "2147483629"):
        assert not re.search(
            rf"(?<![0-9A-Za-z_.]){literal}(?![0-9A-Za-z_.])", text
        ), f"hash constant {literal} appears in the calculation contract"


def test_the_hash_base_cannot_be_added_to_the_contract() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["fingerprint"]["fp_base"] = 131

    _rejected(mutate, "FP_BASE copied into the calculation contract")


def test_a_modulus_cannot_be_added_to_the_contract() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["fingerprint"]["fp_mod_1"] = 2147483647

    _rejected(mutate, "a hash modulus copied into the calculation contract")


def test_a_hash_constant_hidden_under_an_innocent_key_is_still_caught() -> None:
    """The literal scan is what catches a renamed copy."""

    def mutate(data: dict[str, Any]) -> None:
        data["fingerprint"]["note"] = "the second modulus is 2147483629"

    _rejected(mutate, "a hash modulus smuggled in as prose")


def test_a_recurrence_key_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["fingerprint"]["recurrence"] = "h = h * b + u"

    _rejected(mutate, "the hash recurrence restated in the contract")


def test_the_contract_does_not_restate_the_distribution_master_list() -> None:
    """Distribution authority belongs to input_contract.yaml.

    The scan is over the PARSED data, not the file text: a comment naming the
    owning authority is documentation of the boundary, whereas a value would be a
    second copy that could drift.
    """
    data = _base()
    found: list[str] = []

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            for value in node.values():
                walk(value)
        elif isinstance(node, list):
            for value in node:
                walk(value)
        elif isinstance(node, str):
            for name in ("Triangular", "Beta-PERT", "Uniform"):
                if name in node:
                    found.append(node)

    walk(data)
    assert not found, f"the calculation contract restates distribution names: {found}"


def test_every_authority_reference_resolves_in_its_owner() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert calc.authority_references
    validate_calc_against(
        calc,
        load_spec(SPEC_PATH),
        load_contract(CONTRACT_PATH),
        load_driver_contract(DRIVERS_PATH),
        load_structure_contract(STRUCTURE_PATH),
    )


def test_a_dangling_authority_reference_fails_cross_validation() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["authority_references"][0]["locator"] = "config_tables.no_such_table"

    _rejected_cross(mutate, "an authority reference that no longer resolves")


def test_an_authority_reference_to_a_stranger_file_fails_cross_validation() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["authority_references"][0]["owner"] = "somewhere_else.yaml"

    _rejected_cross(mutate, "an authority reference outside the four accepted contracts")


def test_a_table_name_colliding_with_another_contract_fails_cross_validation() -> None:
    """Defence in depth, reached by replacement rather than by YAML.

    The load-time check locks the five table names, so a colliding name cannot get
    that far through the file. The cross-check answers a different question - is
    this name already taken by one of the other four contracts? - and is asserted
    here on its own.
    """
    calc = load_calc_contract(CALC_PATH)
    tables = dict(calc.tables)
    tables["calc_fx"] = dataclasses.replace(tables["calc_fx"], table_name="tblCostLines")
    _rejected_cross_replace(
        tables=tables, reason="a Phase-5 table name already owned by another contract"
    )


# ---------------------------------------------------------------------------
# the sheet
# ---------------------------------------------------------------------------
def test_the_calc_sheet_is_hidden_not_very_hidden() -> None:
    """`veryHidden` would put the calculation record beyond an auditor's reach."""
    calc = load_calc_contract(CALC_PATH)
    assert calc.required_visibility == "hidden"
    sheet = next(s for s in load_spec(SPEC_PATH).sheets if s.name == "_Calc")
    assert sheet.visibility == "hidden"


def test_a_very_hidden_calc_sheet_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["sheet"]["required_visibility"] = "veryHidden"

    _rejected(mutate, "a veryHidden calculation sheet")


def test_a_visibility_disagreement_with_the_manifest_fails_cross_validation() -> None:
    """Reached by replacement: the load-time rule already refuses anything but
    'hidden', so the cross-check is only exercisable on its own terms."""
    _rejected_cross_replace(
        required_visibility="visible",
        reason="a required visibility that disagrees with workbook.yaml",
    )


def test_phase5_blocks_are_placed_on_the_calc_sheet() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["sheet"]["name"] = "Setup"

    _rejected(mutate, "Phase-5 blocks placed on a user-facing sheet")


# ---------------------------------------------------------------------------
# tolerances
# ---------------------------------------------------------------------------
def test_the_locked_tolerance_constants() -> None:
    tolerances = load_calc_contract(CALC_PATH).tolerances
    assert tolerances.profiling_sum_absolute == 1e-9
    assert tolerances.identity_absolute_floor == 1e-6
    assert tolerances.identity_relative_coefficient == 1e-12
    assert tolerances.conditioning_scale_floor == 1.0
    assert tolerances.fx_rate_strictly_positive is True
    assert tolerances.growth_factor_strictly_positive is True


def test_every_identity_declares_a_cancellation_safe_conditioning_scale() -> None:
    """The scale must sum ABSOLUTE terms, or a cancelling model collapses its own
    tolerance to the floor and reports ordinary accumulation error as a mismatch."""
    tolerances = load_calc_contract(CALC_PATH).tolerances
    assert set(tolerances.conditioning_terms) == {
        "i1",
        "i2",
        "i3a",
        "i3b",
        "i3c",
        "i4a",
        "i4b",
        "i4c",
    }
    for identity, terms in tolerances.conditioning_terms.items():
        assert len(terms) >= 2, identity
        assert all(term.startswith(("abs_", "sum_abs_")) for term in terms), identity


def test_a_single_term_conditioning_scale_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["tolerances"]["conditioning_terms"]["i1"] = ["abs_c"]

    _rejected(mutate, "a conditioning scale driven by the net result alone")


def test_a_conditioning_floor_below_unity_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["tolerances"]["conditioning_scale_floor"] = 0.0

    _rejected(mutate, "a conditioning floor that could shrink the tolerance")


def test_a_relaxed_fx_positivity_rule_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["tolerances"]["fx_rate_strictly_positive"] = False

    _rejected(mutate, "FX positivity relaxed to allow an epsilon")


def test_a_negative_tolerance_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["tolerances"]["profiling_sum_absolute"] = -1e-9

    _rejected(mutate, "a negative tolerance")


# ---------------------------------------------------------------------------
# structural malformation
# ---------------------------------------------------------------------------
def test_a_contract_that_is_not_a_mapping_is_rejected() -> None:
    with tempfile.TemporaryDirectory(prefix="pccm-badcalc-") as tmp:
        path = Path(tmp) / "list.yaml"
        path.write_text("- not a mapping\n", encoding="utf-8")
        try:
            load_calc_contract(path)
        except CalcContractError:
            return
    raise AssertionError("a non-mapping calculation contract was silently accepted")


def test_a_missing_top_level_section_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        del data["tolerances"]

    _rejected(mutate, "a missing top-level section")


def test_an_unknown_value_type_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _table(data, "tblCalcYears")["columns"][0]["value_type"] = "money"

    _rejected(mutate, "an unknown column value type")


def test_a_non_pascal_table_name_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _table(data, "tblCalcFX")["table_name"] = "calc_fx"

    _rejected(mutate, "a table name outside the tbl<PascalCase> convention")


# ---------------------------------------------------------------------------
# runner
# ---------------------------------------------------------------------------
def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 5 Gate-A Step-1 calculation contract negative tests")
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
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(_run_all())
