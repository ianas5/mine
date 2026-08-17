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
    LOCKED_AUTHORITY_REFERENCES,
    LOCKED_CALC_CONTRACT_VERSION,
    LOCKED_CONDITIONING_TERMS,
    LOCKED_TABLE_KEYS,
    LOCKED_TABLES,
    LOCKED_TOLERANCES,
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


_BASE_CACHE: dict[str, Any] | None = None


def _base() -> dict[str, Any]:
    """A fresh copy of the real contract, parsed once.

    The per-attribute sweeps mutate several hundred copies; re-parsing the YAML for
    each one dominated the suite's runtime and proved nothing extra.
    """
    global _BASE_CACHE
    if _BASE_CACHE is None:
        with CALC_PATH.open("r", encoding="utf-8") as handle:
            _BASE_CACHE = yaml.safe_load(handle)
    return copy.deepcopy(_BASE_CACHE)


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


def _totals_field(data: dict[str, Any], key: str) -> dict[str, Any]:
    return next(f for f in data["scalar_blocks"]["calc_totals"]["fields"] if f["key"] == key)


def _table(data: dict[str, Any], table_name: str) -> dict[str, Any]:
    return next(t for t in data["tables"].values() if t["table_name"] == table_name)


def _column(data: dict[str, Any], table_name: str, key: str) -> dict[str, Any]:
    return next(c for c in _table(data, table_name)["columns"] if c["key"] == key)


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


def test_a_dangling_authority_reference_is_rejected() -> None:
    """Caught twice: the set lock at load time, the resolver at cross-validation."""

    def mutate(data: dict[str, Any]) -> None:
        data["authority_references"][0]["locator"] = "config_tables.no_such_table"

    _rejected(mutate, "an authority reference that no longer resolves")


def test_a_dangling_locator_still_fails_the_resolver_on_its_own() -> None:
    """The load-time set lock now fires first, so the resolver is reached by
    replacement. It must still refuse a locator that does not resolve."""
    calc = load_calc_contract(CALC_PATH)
    references = list(calc.authority_references)
    references[0] = dataclasses.replace(references[0], locator="config_tables.no_such_table")
    _rejected_cross_replace(
        authority_references=tuple(references),
        reason="a locator that does not resolve in its owning contract",
    )


def test_an_authority_reference_to_a_stranger_file_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["authority_references"][0]["owner"] = "somewhere_else.yaml"

    _rejected(mutate, "an authority reference outside the four accepted contracts")


def test_a_stranger_owner_still_fails_the_resolver_on_its_own() -> None:
    calc = load_calc_contract(CALC_PATH)
    references = list(calc.authority_references)
    references[0] = dataclasses.replace(references[0], owner="somewhere_else.yaml")
    _rejected_cross_replace(
        authority_references=tuple(references),
        reason="an owner outside the four accepted specifications",
    )


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
# BLOCKER 1 regression - the exact mutations independent review demonstrated
#
# Every one of these was ACCEPTED by the loader at commit f6a35fe. Each changes
# the accepted Revision-E design while moving no anchor and breaking no syntactic
# rule, which is precisely why a "looks structurally valid" check cannot see them.
# ---------------------------------------------------------------------------
def test_r1_renaming_a_years_column_header_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _column(data, "tblCalcYears", "project_index")["header"] = "Project Number"

    _rejected(mutate, "tblCalcYears Project Index header renamed")


def test_r2_changing_a_years_column_number_format_is_rejected() -> None:
    """An index shown to two decimal places is not the accepted design."""

    def mutate(data: dict[str, Any]) -> None:
        _column(data, "tblCalcYears", "project_index")["number_format"] = "0.00"

    _rejected(mutate, "tblCalcYears Project Index number_format changed")


def test_r3_renaming_the_fx_rate_header_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _column(data, "tblCalcFX", "fx_to_sar")["header"] = "Rate"

    _rejected(mutate, "tblCalcFX 'FX to SAR' header renamed")


def test_r4_changing_the_fx_units_is_rejected() -> None:
    """`SAR per unit` IS the FX convention made visible in the audit table."""

    def mutate(data: dict[str, Any]) -> None:
        _column(data, "tblCalcFX", "fx_to_sar")["units"] = "USD"

    _rejected(mutate, "tblCalcFX units changed to USD")


def test_r5_downgrading_a_timestamp_to_text_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        _state_field(data, "last_successful_stamp")["value_type"] = "text"

    _rejected(mutate, "calc_state Last Successful Stamp value_type downgraded to text")


def test_r6_replacing_a_timestamp_format_with_at_is_rejected() -> None:
    """A stamp stored as `@` is unsortable and uncomparable."""

    def mutate(data: dict[str, Any]) -> None:
        _state_field(data, "last_successful_stamp")["number_format"] = "@"

    _rejected(mutate, "calc_state Last Successful Stamp number_format replaced with @")


def test_r7_relabelling_a_headline_total_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["scalar_blocks"]["calc_totals"]["fields"][0]["label"] = "Some Other Total"

    _rejected(mutate, "calc_totals first label replaced with arbitrary text")


def test_r8_changing_a_headline_total_number_format_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["scalar_blocks"]["calc_totals"]["fields"][0]["number_format"] = "0"

    _rejected(mutate, "calc_totals first number_format changed to 0")


def test_r9_widening_quantity_to_risk_rows_is_rejected() -> None:
    """Quantity is a Cost Line field. Letting a Risk row carry one would make the
    column mean two different things by kind - the exact rule §16.4 forbids."""

    def mutate(data: dict[str, Any]) -> None:
        _column(data, "tblCalcDrivers", "quantity")["applies_to"] = ["cost_line", "risk"]

    _rejected(mutate, "tblCalcDrivers Quantity widened to risk rows")


def test_r10_widening_expected_risk_to_cost_lines_is_rejected() -> None:
    """`D_nom = SUM(Expected Risk Nominal) over RISK rows`. Widening the column
    silently changes what that sum means."""

    def mutate(data: dict[str, Any]) -> None:
        _column(data, "tblCalcDrivers", "expected_risk_nominal")["applies_to"] = [
            "cost_line",
            "risk",
        ]

    _rejected(mutate, "tblCalcDrivers Expected Risk Nominal widened to cost lines")


def test_r11_narrowing_knom_away_from_risk_rows_is_rejected() -> None:
    """Knom applies to both kinds: a Risk's expected value is escalated too."""

    def mutate(data: dict[str, Any]) -> None:
        _column(data, "tblCalcDrivers", "knom")["applies_to"] = ["cost_line"]

    _rejected(mutate, "tblCalcDrivers Knom narrowed away from risk rows")


def test_r12_loosening_the_profiling_tolerance_is_rejected() -> None:
    """1e-3 would let a profile summing to 99.9% pass as 100%."""

    def mutate(data: dict[str, Any]) -> None:
        data["tolerances"]["profiling_sum_absolute"] = 1e-3

    _rejected(mutate, "profiling_sum_absolute loosened from 1e-9 to 1e-3")


def test_r13_loosening_the_identity_floor_is_rejected() -> None:
    """1e-3 SAR would let a real bookkeeping mismatch pass as rounding."""

    def mutate(data: dict[str, Any]) -> None:
        data["tolerances"]["identity_absolute_floor"] = 1e-3

    _rejected(mutate, "identity_absolute_floor loosened from 1e-6 to 1e-3")


def test_r14_raising_the_conditioning_floor_is_rejected() -> None:
    """A floor of 10 widens every identity tolerance at once."""

    def mutate(data: dict[str, Any]) -> None:
        data["tolerances"]["conditioning_scale_floor"] = 10

    _rejected(mutate, "conditioning_scale_floor raised from 1.0 to 10")


def test_r15_giving_an_identity_the_wrong_conditioning_terms_is_rejected() -> None:
    """I1 is `A + B = C`; sizing its tolerance by |D| and |E| is meaningless."""

    def mutate(data: dict[str, Any]) -> None:
        data["tolerances"]["conditioning_terms"]["i1"] = ["abs_d", "abs_e"]

    _rejected(mutate, "conditioning_terms.i1 replaced with another identity's terms")


# ---------------------------------------------------------------------------
# BLOCKER 1 - exhaustive per-attribute sweep
#
# The fifteen cases above are the ones review happened to try. This sweep is the
# general guard: EVERY attribute of EVERY locked column and EVERY locked scalar
# row must be individually load-bearing.
# ---------------------------------------------------------------------------
def _altered(attribute: str, current: Any) -> Any:
    """A different, still-syntactically-plausible value for one attribute."""
    if attribute == "key":
        return f"{current}_x"
    if attribute == "header":
        return f"{current} X"
    if attribute == "label":
        return f"{current} (revised)"
    if attribute == "value_type":
        return "text" if current != "text" else "double"
    if attribute == "number_format":
        return "0.0000" if current != "0.0000" else "0.00"
    if attribute == "units":
        return "widgets"
    if attribute == "measure":
        return "Z"
    if attribute == "basis":
        return "nominal" if current != "nominal" else "pv"
    if attribute == "enum":
        return "derived_status" if current != "derived_status" else "attempt_result"
    if attribute == "initial":
        return None if current is not None else "SEEDED"
    if attribute == "applies_to":
        options = [["cost_line"], ["risk"], ["cost_line", "risk"]]
        return next(o for o in options if o != list(current or []))
    if attribute == "row_rule":
        return f"{current} (revised)"
    raise AssertionError(f"no alteration defined for {attribute!r}")


def _sweep(targets: list[tuple[str, Callable[[dict[str, Any]], dict[str, Any]], list[str]]]) -> None:
    """Every listed attribute of every listed entry must be individually locked."""
    accepted: list[str] = []
    for label, locate, attributes in targets:
        for attribute in attributes:
            data = copy.deepcopy(_base())
            entry = locate(data)
            if attribute not in entry:
                continue
            entry[attribute] = _altered(attribute, entry[attribute])
            with tempfile.TemporaryDirectory(prefix="pccm-sweep-") as tmp:
                try:
                    load_calc_contract(_write(data, tmp))
                except CalcContractError:
                    continue
                except Exception as error:  # noqa: BLE001
                    raise AssertionError(
                        f"{label}.{attribute}: raised {type(error).__name__} instead of "
                        "CalcContractError"
                    ) from error
            accepted.append(f"{label}.{attribute}")
    assert not accepted, "silently accepted design changes: " + ", ".join(accepted)


COLUMN_ATTRIBUTES = ["key", "header", "value_type", "number_format", "units", "applies_to"]


def test_every_attribute_of_every_table_column_is_locked() -> None:
    calc = load_calc_contract(CALC_PATH)
    targets = []
    for table in calc.all_tables:
        for column in table.columns:
            targets.append(
                (
                    f"{table.table_name}.{column.key}",
                    (
                        lambda data, t=table.table_name, k=column.key: _column(data, t, k)
                    ),
                    COLUMN_ATTRIBUTES,
                )
            )
    assert len(targets) == 39, f"expected 39 locked columns across five tables, got {len(targets)}"
    _sweep(targets)


def test_every_table_row_rule_is_locked() -> None:
    calc = load_calc_contract(CALC_PATH)
    _sweep(
        [
            (t.table_name, (lambda data, n=t.table_name: _table(data, n)), ["row_rule"])
            for t in calc.all_tables
        ]
    )


def test_every_attribute_of_every_calc_state_row_is_locked() -> None:
    calc = load_calc_contract(CALC_PATH)
    _sweep(
        [
            (
                f"calc_state.{f.key}",
                (lambda data, k=f.key: _state_field(data, k)),
                ["key", "label", "value_type", "number_format", "enum", "initial"],
            )
            for f in calc.calc_state.fields
        ]
    )


def test_every_attribute_of_every_calc_totals_row_is_locked() -> None:
    calc = load_calc_contract(CALC_PATH)
    _sweep(
        [
            (
                f"calc_totals.{f.key}",
                (lambda data, k=f.key: _totals_field(data, k)),
                ["key", "label", "value_type", "number_format", "units", "measure", "basis"],
            )
            for f in calc.calc_totals.fields
        ]
    )


def test_an_applies_to_added_to_a_non_driver_table_is_rejected() -> None:
    """`applies_to` is meaningful only on the driver audit table."""

    def mutate(data: dict[str, Any]) -> None:
        _column(data, "tblCalcAnnual", "total_pv")["applies_to"] = ["cost_line"]

    _rejected(mutate, "applies_to declared on a table that has no driver kinds")


def test_the_locked_design_matches_the_contract_exactly() -> None:
    """The loader's copy and the contract agree today - the positive direction."""
    calc = load_calc_contract(CALC_PATH)
    for name, schema in LOCKED_TABLES.items():
        table = calc.table_by_name(name)
        assert table.row_rule == schema.row_rule
        assert len(table.columns) == len(schema.columns)
        for expected, got in zip(schema.columns, table.columns):
            assert (got.key, got.header, got.value_type, got.number_format, got.units) == (
                expected.key,
                expected.header,
                expected.value_type,
                expected.number_format,
                expected.units,
            )
            assert got.applies_to == expected.applies_to


# ---------------------------------------------------------------------------
# BLOCKER 3 regression - the required authority-reference set is complete
# ---------------------------------------------------------------------------
def test_the_six_required_authority_references_are_declared() -> None:
    calc = load_calc_contract(CALC_PATH)
    declared = tuple((r.concept, r.owner, r.locator) for r in calc.authority_references)
    assert declared == LOCKED_AUTHORITY_REFERENCES
    assert len(declared) == 6


def test_removing_the_fx_convention_reference_is_rejected() -> None:
    """Accepted at f6a35fe: the boundary simply stopped being declared."""

    def mutate(data: dict[str, Any]) -> None:
        data["authority_references"] = [
            r for r in data["authority_references"] if r["concept"] != "FX convention"
        ]

    _rejected(mutate, "the FX convention authority reference removed")


def test_removing_the_distribution_reference_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["authority_references"] = [
            r
            for r in data["authority_references"]
            if r["concept"] != "distribution master list"
        ]

    _rejected(mutate, "the distribution authority reference removed")


def test_removing_any_single_authority_reference_is_rejected() -> None:
    for index in range(len(LOCKED_AUTHORITY_REFERENCES)):

        def mutate(data: dict[str, Any], i: int = index) -> None:
            del data["authority_references"][i]

        _rejected(mutate, f"authority reference {index} removed")


def test_duplicating_an_authority_reference_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["authority_references"].append(copy.deepcopy(data["authority_references"][0]))

    _rejected(mutate, "a duplicated authority reference")


def test_a_concept_declared_twice_with_different_owners_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        clone = copy.deepcopy(data["authority_references"][0])
        clone["owner"] = "workbook.yaml"
        clone["locator"] = "sheets"
        data["authority_references"].append(clone)

    _rejected(mutate, "one concept claimed by two owners")


def test_changing_an_authority_owner_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        entry = next(r for r in data["authority_references"] if r["concept"] == "FX convention")
        entry["owner"] = "structure_contract.yaml"

    _rejected(mutate, "an authority reference pointed at a different owner")


def test_changing_an_authority_locator_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        entry = next(r for r in data["authority_references"] if r["concept"] == "FX convention")
        entry["locator"] = "conventions.input_prefix"

    _rejected(mutate, "an authority reference redirected to a different locator")


def test_renaming_a_concept_is_rejected() -> None:
    """A rename looks like a documentation tidy-up and silently replaces a
    boundary: the set would still have six entries."""

    def mutate(data: dict[str, Any]) -> None:
        entry = next(r for r in data["authority_references"] if r["concept"] == "FX convention")
        entry["concept"] = "currency handling"

    _rejected(mutate, "a renamed authority concept")


def test_an_unexpected_authority_reference_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["authority_references"].append(
            {
                "concept": "something else entirely",
                "owner": "workbook.yaml",
                "locator": "presentation",
            }
        )

    _rejected(mutate, "an unexpected authority reference")


# ---------------------------------------------------------------------------
# REVISION-E DESIGN PARITY
#
# A THIRD, INDEPENDENT COPY of the accepted design, transcribed by hand from
# docs/phase5_plan.md sections 16.3-16.4 and 15.
#
# It deliberately does NOT read LOCKED_TABLES, LOCKED_CALC_STATE,
# LOCKED_CALC_TOTALS, LOCKED_TOLERANCES or LOCKED_TABLE_KEYS. A guard built by
# copying the first YAML implementation only proves the YAML matches itself: the
# contract and the lock could drift away from the plan together and still confirm
# each other. These literals are what the plan says, and BOTH the contract and the
# loader's copy are checked against them.
# ---------------------------------------------------------------------------
EM_DASH = "—"
"""The plan's own notation for "no unit applies". Not "name", not "key" - a
categorical identifier has no physical unit, and inventing a pseudo-unit invents a
business meaning the plan does not have."""

PLAN_TABLE_KEYS = {
    "calc_years": "tblCalcYears",
    "calc_inflation_factors": "tblCalcInflationFactors",
    "calc_fx": "tblCalcFX",
    "calc_drivers": "tblCalcDrivers",
    "calc_annual": "tblCalcAnnual",
}

# (header, value_type, number_format, units)
PLAN_YEARS = (
    ("Project Index", "integer", "0", "index, from 1"),
    ("Calendar Year", "integer", "0", "year"),
    ("Discount Factor", "double", "0.000000", "dimensionless"),
)
PLAN_INFLATION = (
    ("Inflation Profile", "text", "@", "key"),
    ("Calendar Year", "integer", "0", "year"),
    ("Annual Rate", "double", "0.00%", "rate"),
    ("Cumulative Inflation Factor", "double", "0.000000", "dimensionless"),
)
PLAN_FX = (
    ("Currency", "text", "@", "key"),
    ("FX to SAR", "double", "0.000000", "SAR per unit"),
    ("Referenced By", "integer", "0", "driver count"),
)
PLAN_DRIVERS = (
    ("Permanent ID", "text", "@", "key"),
    ("Driver Kind", "text", "@", "Cost Line / Risk"),
    ("Distribution", "text", "@", EM_DASH),
    ("Central Basis", "text", "@", "ML / Midpoint"),
    ("Currency", "text", "@", EM_DASH),
    ("FX to SAR", "double", "0.000000", "SAR per unit"),
    ("Inflation Profile", "text", "@", EM_DASH),
    ("Quantity", "double", "#,##0.00", "units"),
    ("Probability", "double", "0.0%", "fraction"),
    ("Central Value", "double", "#,##0.00", "source currency"),
    ("Mean Value", "double", "#,##0.00", "source currency"),
    ("Knom", "double", "0.000000", "SAR per source unit"),
    ("Kpv", "double", "0.000000", "SAR per source unit"),
    ("Deterministic Nominal", "double", "#,##0.00", "SAR"),
    ("Deterministic PV", "double", "#,##0.00", "SAR"),
    ("Mean-Basis Nominal", "double", "#,##0.00", "SAR"),
    ("Mean-Basis PV", "double", "#,##0.00", "SAR"),
    ("Uncertainty Mean Shift Nominal", "double", "#,##0.00", "SAR"),
    ("Uncertainty Mean Shift PV", "double", "#,##0.00", "SAR"),
    ("Expected Risk Nominal", "double", "#,##0.00", "SAR"),
    ("Expected Risk PV", "double", "#,##0.00", "SAR"),
)
PLAN_ANNUAL = (
    ("Project Index", "integer", "0", "index"),
    ("Calendar Year", "integer", "0", "year"),
    ("Base Cost Nominal", "double", "#,##0.00", "SAR"),
    ("Expected Risk Nominal", "double", "#,##0.00", "SAR"),
    ("Total Nominal", "double", "#,##0.00", "SAR"),
    ("Base Cost PV", "double", "#,##0.00", "SAR"),
    ("Expected Risk PV", "double", "#,##0.00", "SAR"),
    ("Total PV", "double", "#,##0.00", "SAR"),
)
PLAN_SCHEMAS = {
    "tblCalcYears": PLAN_YEARS,
    "tblCalcInflationFactors": PLAN_INFLATION,
    "tblCalcFX": PLAN_FX,
    "tblCalcDrivers": PLAN_DRIVERS,
    "tblCalcAnnual": PLAN_ANNUAL,
}

# Which driver kinds populate each tblCalcDrivers column, from the plan's
# "Cost Line" / "Risk" columns: "blank" means the kind does not populate it.
PLAN_DRIVER_APPLIES = (
    ("Permanent ID", ("cost_line", "risk")),
    ("Driver Kind", ("cost_line", "risk")),
    ("Distribution", ("cost_line", "risk")),
    ("Central Basis", ("cost_line", "risk")),
    ("Currency", ("cost_line", "risk")),
    ("FX to SAR", ("cost_line", "risk")),
    ("Inflation Profile", ("cost_line", "risk")),
    ("Quantity", ("cost_line",)),
    ("Probability", ("risk",)),
    ("Central Value", ("cost_line",)),
    ("Mean Value", ("cost_line", "risk")),
    ("Knom", ("cost_line", "risk")),
    ("Kpv", ("cost_line", "risk")),
    ("Deterministic Nominal", ("cost_line",)),
    ("Deterministic PV", ("cost_line",)),
    ("Mean-Basis Nominal", ("cost_line",)),
    ("Mean-Basis PV", ("cost_line",)),
    ("Uncertainty Mean Shift Nominal", ("cost_line",)),
    ("Uncertainty Mean Shift PV", ("cost_line",)),
    ("Expected Risk Nominal", ("risk",)),
    ("Expected Risk PV", ("risk",)),
)

# (row, group, label, value_type, number_format, initial)
PLAN_CALC_STATE = (
    (13, "snapshot", "Last Successful Stamp", "timestamp", "yyyy-mm-dd hh:mm:ss", None),
    (14, "snapshot", "Last Successful Fingerprint", "text", "@", None),
    (15, "snapshot", "Fingerprint Version", "integer", "0", None),
    (16, "snapshot", "Last Successful Applied Timeline", "text", "@", None),
    (17, "attempt", "Last Attempt Result", "enum", "@", "NONE"),
    (18, "attempt", "Last Attempt Detail", "text", "@", None),
    (19, "derived", "Calculation Status (last evaluated)", "enum", "@", "NOT CALCULATED"),
    (20, "derived", "Status Evaluated At", "timestamp", "yyyy-mm-dd hh:mm:ss", None),
)

# (row, measure, basis, label) - all `#,##0.00`, all SAR, none seeded.
PLAN_CALC_TOTALS = (
    (23, "A", "nominal", f"Escalated Deterministic Base {EM_DASH} Nominal"),
    (24, "A", "pv", f"Escalated Deterministic Base {EM_DASH} PV"),
    (25, "B", "nominal", f"Uncertainty Mean Shift {EM_DASH} Nominal"),
    (26, "B", "pv", f"Uncertainty Mean Shift {EM_DASH} PV"),
    (27, "C", "nominal", f"Mean-Basis Base Cost {EM_DASH} Nominal"),
    (28, "C", "pv", f"Mean-Basis Base Cost {EM_DASH} PV"),
    (29, "D", "nominal", f"Expected Risk / EMV {EM_DASH} Nominal"),
    (30, "D", "pv", f"Expected Risk / EMV {EM_DASH} PV"),
    (31, "E", "nominal", f"Analytical Mean Total {EM_DASH} Nominal"),
    (32, "E", "pv", f"Analytical Mean Total {EM_DASH} PV"),
)

PLAN_TOLERANCES = {
    "profiling_sum_absolute": 1e-9,
    "identity_absolute_floor": 1e-6,
    "identity_relative_coefficient": 1e-12,
    "conditioning_scale_floor": 1.0,
    "fx_rate_strictly_positive": True,
    "growth_factor_strictly_positive": True,
}

PLAN_ANCHORS = {
    "tblCalcYears": ("H", "J", 15),
    "tblCalcInflationFactors": ("M", "P", 15),
    "tblCalcFX": ("S", "U", 15),
    "tblCalcDrivers": ("X", "AR", 15),
    "tblCalcAnnual": ("AU", "BB", 15),
}


def test_parity_the_contract_matches_the_accepted_plan_table_schemas() -> None:
    calc = load_calc_contract(CALC_PATH)
    for table_name, expected in PLAN_SCHEMAS.items():
        table = calc.table_by_name(table_name)
        actual = tuple(
            (c.header, c.value_type, c.number_format, c.units) for c in table.columns
        )
        assert actual == expected, table_name


def test_parity_the_loader_lock_matches_the_accepted_plan_table_schemas() -> None:
    """The independent guard must encode Revision E, not the first YAML draft."""
    for table_name, expected in PLAN_SCHEMAS.items():
        schema = LOCKED_TABLES[table_name]
        actual = tuple(
            (c.header, c.value_type, c.number_format, c.units) for c in schema.columns
        )
        assert actual == expected, table_name
        first, last, header_row = PLAN_ANCHORS[table_name]
        assert (schema.first_column, schema.last_column, schema.header_row) == (
            first,
            last,
            header_row,
        )


def test_parity_categorical_driver_columns_carry_no_unit() -> None:
    """Distribution, Currency and Inflation Profile are categorical identifiers.

    The plan records their unit as an em dash. Earlier drafts wrote "name" and
    "key", which read as units and are not.
    """
    calc = load_calc_contract(CALC_PATH)
    drivers = calc.table_by_name("tblCalcDrivers")
    for key in ("distribution", "currency", "inflation_profile"):
        column = next(c for c in drivers.columns if c.key == key)
        assert column.units == EM_DASH, key
        assert column.units not in ("name", "key")


def test_parity_driver_applies_to_matches_the_plan_kind_columns() -> None:
    calc = load_calc_contract(CALC_PATH)
    drivers = calc.table_by_name("tblCalcDrivers")
    actual = tuple((c.header, c.applies_to) for c in drivers.columns)
    assert actual == PLAN_DRIVER_APPLIES


def test_parity_the_contract_matches_the_accepted_plan_calc_state() -> None:
    calc = load_calc_contract(CALC_PATH)
    actual = tuple(
        (f.row, f.group, f.label, f.value_type, f.number_format, f.initial)
        for f in calc.calc_state.fields
    )
    assert actual == PLAN_CALC_STATE


def test_parity_the_contract_matches_the_accepted_plan_calc_totals() -> None:
    calc = load_calc_contract(CALC_PATH)
    actual = tuple((f.row, f.measure, f.basis, f.label) for f in calc.calc_totals.fields)
    assert actual == PLAN_CALC_TOTALS
    for field in calc.calc_totals.fields:
        assert field.value_type == "double"
        assert field.number_format == "#,##0.00"
        assert field.units == "SAR"
        assert field.initial is None


def test_parity_headline_labels_use_the_plan_em_dash() -> None:
    """The contract owns labels, and the plan's separator is an em dash."""
    calc = load_calc_contract(CALC_PATH)
    for field in calc.calc_totals.fields:
        assert EM_DASH in field.label, field.key
        assert f" - {field.basis.upper()}" not in field.label.upper()
    # "Mean-Basis" keeps its ordinary hyphen; only the separator is an em dash.
    c_nom = calc.calc_totals.field_by_key("c_nom")
    assert c_nom.label == f"Mean-Basis Base Cost {EM_DASH} Nominal"
    assert "Mean-Basis" in c_nom.label


def test_parity_the_table_mapping_keys_match_the_plan() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert {key: table.table_name for key, table in calc.tables.items()} == PLAN_TABLE_KEYS


def test_parity_the_tolerances_match_the_plan() -> None:
    tolerances = load_calc_contract(CALC_PATH).tolerances
    for name, expected in PLAN_TOLERANCES.items():
        assert getattr(tolerances, name) == expected, name


# ---------------------------------------------------------------------------
# the contract document's own identity
# ---------------------------------------------------------------------------
def test_the_contract_declares_the_supported_document_version() -> None:
    assert load_calc_contract(CALC_PATH).version == "1.0.0"
    assert LOCKED_CALC_CONTRACT_VERSION == "1.0.0"


def test_an_unsupported_contract_version_is_rejected() -> None:
    """A loader for 1.0.0 must not silently consume another format.

    All three were accepted at commit d0c9bca.
    """
    for version in ("1.0.1", "2.0.0", "0.9.0", "9.9.9", "foo", "1.0", ""):

        def mutate(data: dict[str, Any], v: str = version) -> None:
            data["calc_contract_version"] = v

        _rejected(mutate, f"contract version {version!r}")


def test_the_contract_version_and_fp_version_are_separate_domains() -> None:
    """Different questions: which document format, versus which digest encoding."""
    calc = load_calc_contract(CALC_PATH)
    assert calc.version == "1.0.0"
    assert calc.fingerprint_version == 1
    assert LOCKED_CALC_CONTRACT_VERSION != str(LOCKED_FP_VERSION)


# ---------------------------------------------------------------------------
# the table mapping keys are part of the contract
# ---------------------------------------------------------------------------
def test_the_five_table_mapping_keys_are_locked() -> None:
    calc = load_calc_contract(CALC_PATH)
    assert set(calc.tables) == set(LOCKED_TABLE_KEYS) == set(PLAN_TABLE_KEYS)
    assert LOCKED_TABLE_KEYS == PLAN_TABLE_KEYS


def test_renaming_any_table_mapping_key_is_rejected() -> None:
    """Accepted at d0c9bca: `calc_fx` -> `foo` left the ListObject name intact, so
    every consumer addressing the block semantically would break silently."""
    for key in PLAN_TABLE_KEYS:

        def mutate(data: dict[str, Any], k: str = key) -> None:
            data["tables"][f"{k}_renamed"] = data["tables"].pop(k)

        _rejected(mutate, f"table mapping key {key!r} renamed")


def test_the_two_specific_renames_review_demonstrated_are_rejected() -> None:
    def rename_fx(data: dict[str, Any]) -> None:
        data["tables"]["foo"] = data["tables"].pop("calc_fx")

    def rename_years(data: dict[str, Any]) -> None:
        data["tables"]["years2"] = data["tables"].pop("calc_years")

    _rejected(rename_fx, "tables.calc_fx renamed to foo")
    _rejected(rename_years, "tables.calc_years renamed to years2")


def test_a_missing_table_mapping_key_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        del data["tables"]["calc_drivers"]

    _rejected(mutate, "a table mapping key removed")


def test_a_semantic_key_pointed_at_the_wrong_listobject_is_rejected() -> None:
    """Both halves of the pair are locked, and so is the pairing."""

    def mutate(data: dict[str, Any]) -> None:
        data["tables"]["calc_fx"]["table_name"] = "tblCalcAnnual"
        data["tables"]["calc_annual"]["table_name"] = "tblCalcFX"

    _rejected(mutate, "two semantic keys swapped onto each other's ListObjects")


# ---------------------------------------------------------------------------
# no silently ignored machine-readable fields
# ---------------------------------------------------------------------------
REMOVED_REDUNDANT_FIELDS = (
    ("phase4_reservation.owning_contract", ["phase4_reservation"], "owning_contract"),
    ("fingerprint.version_written_by", ["fingerprint"], "version_written_by"),
    ("fingerprint.mathematics_owner", ["fingerprint"], "mathematics_owner"),
    ("calc_state.commit_range", ["scalar_blocks", "calc_state"], "commit_range"),
    ("calc_state.snapshot_range", ["scalar_blocks", "calc_state"], "snapshot_range"),
    ("calc_state.attempt_range", ["scalar_blocks", "calc_state"], "attempt_range"),
    ("calc_state.derived_range", ["scalar_blocks", "calc_state"], "derived_range"),
    ("calc_totals.value_range", ["scalar_blocks", "calc_totals"], "value_range"),
    ("calc_totals.units", ["scalar_blocks", "calc_totals"], "units"),
    ("calc_fx.fx_convention_owner", ["tables", "calc_fx"], "fx_convention_owner"),
)
"""Machine-readable fields that were present, unparsed and unvalidated at d0c9bca.

Every one could be changed to contradict the locked design with the loader
reporting success. Each is now REMOVED from the YAML - the explanation survives as
a comment, which cannot be mistaken for data - and reintroducing any of them is
refused as an unsupported key."""


def test_the_redundant_machine_readable_fields_are_gone_from_the_contract() -> None:
    data = _base()
    for label, trail, key in REMOVED_REDUNDANT_FIELDS:
        node = data
        for step in trail:
            node = node[step]
        assert key not in node, f"{label} is still present in the contract"


def test_reintroducing_any_removed_redundant_field_is_rejected() -> None:
    """The replacement for mutation tests: the fields cannot come back silently."""
    for label, trail, key in REMOVED_REDUNDANT_FIELDS:

        def mutate(data: dict[str, Any], t: list[str] = trail, k: str = key) -> None:
            node = data
            for step in t:
                node = node[step]
            node[k] = "anything at all"

        _rejected(mutate, f"{label} reintroduced")


def test_the_derived_ranges_are_still_available_from_the_parsed_contract() -> None:
    """Removed from YAML because they are DERIVABLE, not because they are unwanted."""
    calc = load_calc_contract(CALC_PATH)
    assert calc.calc_state.value_range() == "C13:C20"
    assert calc.calc_totals.value_range() == "C23:C32"
    snapshot = [f.row for f in calc.calc_state.fields if f.group == "snapshot"]
    attempt = [f.row for f in calc.calc_state.fields if f.group == "attempt"]
    derived = [f.row for f in calc.calc_state.fields if f.group == "derived"]
    assert (snapshot, attempt, derived) == ([13, 14, 15, 16], [17, 18], [19, 20])


# ---------------------------------------------------------------------------
# unknown keys anywhere are refused
# ---------------------------------------------------------------------------
def test_an_unknown_top_level_key_is_rejected() -> None:
    def mutate(data: dict[str, Any]) -> None:
        data["extra_section"] = {"anything": 1}

    _rejected(mutate, "an unknown top-level key")


def test_an_unknown_sheet_key_is_rejected() -> None:
    """`sheet: {foo: bar}` was silently accepted at d0c9bca."""

    def mutate(data: dict[str, Any]) -> None:
        data["sheet"]["foo"] = "bar"

    _rejected(mutate, "sheet.foo")


def test_an_unknown_key_in_every_mapping_level_is_rejected() -> None:
    levels: list[tuple[str, Callable[[dict[str, Any]], dict[str, Any]]]] = [
        ("root", lambda d: d),
        ("sheet", lambda d: d["sheet"]),
        ("phase4_reservation", lambda d: d["phase4_reservation"]),
        ("fingerprint", lambda d: d["fingerprint"]),
        ("state_labels", lambda d: d["state_labels"]),
        ("scalar_blocks.calc_state", lambda d: d["scalar_blocks"]["calc_state"]),
        ("scalar_blocks.calc_totals", lambda d: d["scalar_blocks"]["calc_totals"]),
        ("calc_state field", lambda d: d["scalar_blocks"]["calc_state"]["fields"][0]),
        ("calc_totals field", lambda d: d["scalar_blocks"]["calc_totals"]["fields"][0]),
        ("table", lambda d: d["tables"]["calc_fx"]),
        ("table column", lambda d: d["tables"]["calc_fx"]["columns"][0]),
        ("tolerances", lambda d: d["tolerances"]),
        ("authority reference", lambda d: d["authority_references"][0]),
    ]
    for label, locate in levels:

        def mutate(data: dict[str, Any], where: Any = locate) -> None:
            where(data)["pccm_unknown_key"] = "x"

        _rejected(mutate, f"an unknown key in {label}")


def test_a_typo_in_a_required_key_is_not_silently_dropped() -> None:
    """The practical case: a misspelling leaves the real key missing AND adds an
    unknown one. Either failure is enough; both must be reported, not ignored."""

    def mutate(data: dict[str, Any]) -> None:
        block = data["scalar_blocks"]["calc_state"]
        block["value_colum"] = block.pop("value_column")

    _rejected(mutate, "a misspelled required key")


def test_documentary_notes_remain_permitted_where_they_are_declared() -> None:
    """Notes are an explicit exception, not evidence that anything goes."""
    calc = load_calc_contract(CALC_PATH)
    assert any(f.note for f in calc.calc_state.fields)

    data = _base()
    data["scalar_blocks"]["calc_state"]["fields"][0]["note"] = "a revised explanation"
    with tempfile.TemporaryDirectory(prefix="pccm-note-") as tmp:
        load_calc_contract(_write(data, tmp))   # must NOT raise


# ---------------------------------------------------------------------------
# THE PARSER BOUNDARY - duplicate mapping keys
#
# PyYAML resolves a duplicate key silently, last-one-wins, BEFORE any validator
# runs. Every guard above therefore assumes each field was declared exactly once,
# and that assumption has to be checked at the parser.
#
# These tests work on the RAW CONTRACT TEXT, because `yaml.safe_dump` cannot emit
# a duplicate key: the defect only exists in hand-written YAML, so only
# hand-written YAML can reproduce it.
# ---------------------------------------------------------------------------
def _raw_contract() -> str:
    return CALC_PATH.read_text(encoding="utf-8")


def _rejected_text(text: str, reason: str) -> str:
    """The given contract TEXT must be refused. Returns the error message."""
    with tempfile.TemporaryDirectory(prefix="pccm-dupkey-") as tmp:
        path = Path(tmp) / "duplicated.yaml"
        path.write_text(text, encoding="utf-8")
        try:
            load_calc_contract(path)
        except CalcContractError as error:
            return str(error)
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of CalcContractError"
            ) from error
    raise AssertionError(f"{reason}: silently accepted")


def _duplicate_before(anchor: str, injected: str | None = None) -> str:
    """Insert a competing declaration immediately before the line `anchor`.

    `injected` defaults to a verbatim copy of the anchor line. Passing a DIFFERENT
    value is the important case: two identical values only prove detection exists,
    whereas two different values are the real hazard - a reader sees the first, the
    parser keeps the last.
    """
    text = _raw_contract()
    assert text.count(anchor) >= 1, f"anchor not found in the contract: {anchor!r}"
    index = text.index(anchor)
    duplicate = anchor if injected is None else injected
    return text[:index] + duplicate + text[index:]


def test_the_unmodified_contract_still_loads_under_the_strict_parser() -> None:
    """The guard must not reject the real document."""
    calc = load_calc_contract(CALC_PATH)
    assert calc.version == "1.0.0"
    assert len(calc.all_tables) == 5


def test_a_duplicate_root_key_is_rejected() -> None:
    message = _rejected_text(
        _duplicate_before('calc_contract_version: "1.0.0"\n'),
        "a duplicated root key",
    )
    assert "duplicate key" in message
    assert "calc_contract_version" in message


def test_the_duplicate_key_message_names_the_key_the_file_and_both_lines() -> None:
    message = _rejected_text(
        _duplicate_before('calc_contract_version: "1.0.0"\n'),
        "a duplicated root key",
    )
    assert "duplicate key" in message
    assert "'calc_contract_version'" in message
    assert "duplicated.yaml" in message          # the source file
    assert message.count("line ") >= 2           # both declarations located


def test_a_duplicate_semantic_table_key_is_rejected() -> None:
    """`tables: {calc_fx: ..., calc_fx: ...}` - the second block wins silently."""
    message = _rejected_text(
        _duplicate_before("  calc_fx:\n"), "a duplicated semantic table key"
    )
    assert "'calc_fx'" in message


def test_a_duplicate_nested_table_property_is_rejected() -> None:
    message = _rejected_text(
        _duplicate_before('    table_name: "tblCalcFX"\n'),
        "a duplicated table_name",
    )
    assert "'table_name'" in message


def test_a_contradictory_duplicate_units_declaration_is_rejected() -> None:
    """THE REAL HAZARD: a reader sees `USD`, the validator only ever sees the
    second value, checks it against the locked design, and reports success."""
    message = _rejected_text(
        _duplicate_before('        units: "SAR per unit"\n', '        units: "USD"\n'),
        "contradictory duplicate units",
    )
    assert "'units'" in message


def test_a_contradictory_duplicate_number_format_is_rejected() -> None:
    message = _rejected_text(
        _duplicate_before(
            '        number_format: "yyyy-mm-dd hh:mm:ss"\n',
            '        number_format: "0"\n',
        ),
        "contradictory duplicate number_format",
    )
    assert "'number_format'" in message


def test_a_contradictory_duplicate_tolerance_is_rejected() -> None:
    """A loosened tolerance hidden behind the locked one."""
    message = _rejected_text(
        _duplicate_before(
            "  profiling_sum_absolute: 1.0e-9\n", "  profiling_sum_absolute: 1.0e-3\n"
        ),
        "contradictory duplicate tolerance",
    )
    assert "'profiling_sum_absolute'" in message


def test_a_contradictory_duplicate_conditioning_term_is_rejected() -> None:
    message = _rejected_text(
        _duplicate_before(
            '    i1: ["abs_a", "abs_b", "abs_c"]\n', '    i1: ["abs_d", "abs_e"]\n'
        ),
        "contradictory duplicate conditioning term",
    )
    assert "'i1'" in message


def test_duplicate_keys_are_rejected_at_every_nesting_level() -> None:
    """Recursive by construction: PyYAML calls `construct_mapping` for every
    mapping node, so one guard covers every depth rather than fourteen one-off
    checks."""
    levels: list[tuple[str, str, str | None]] = [
        ("root", 'calc_contract_version: "1.0.0"\n', None),
        ("sheet", '  required_visibility: "hidden"\n', None),
        ("phase4_reservation", "  first_row: 1\n", "  first_row: 2\n"),
        ("fingerprint", "  version: 1\n", "  version: 2\n"),
        ("state_labels", "  derived_status:\n", None),
        ("scalar_blocks", "  calc_state:\n", None),
        ("a scalar block", '    label_column: "B"\n', '    label_column: "D"\n'),
        # NOTE: the anchor for a field or a column must be a key INSIDE its
        # mapping. Duplicating the `- key:` line would start a new list item, not
        # a duplicate mapping key, and would be caught by the schema count instead
        # - proving nothing about this guard.
        ("a scalar field", '        row: 13\n', "        row: 14\n"),
        ("tables", "  calc_years:\n", None),
        ("a table", "    header_row: 15\n", "    header_row: 16\n"),
        ("a table column", '        header: "Project Index"\n', '        header: "Project No"\n'),
        ("tolerances", "  identity_absolute_floor: 1.0e-6\n", "  identity_absolute_floor: 1.0e-3\n"),
        ("conditioning_terms", '    i2: ["abs_c", "abs_d", "abs_e"]\n', None),
        ("an authority reference", '    owner: "input_contract.yaml"\n', None),
    ]
    accepted: list[str] = []
    for label, anchor, injected in levels:
        text = _duplicate_before(anchor, injected)
        with tempfile.TemporaryDirectory(prefix="pccm-dupkey-") as tmp:
            path = Path(tmp) / "duplicated.yaml"
            path.write_text(text, encoding="utf-8")
            try:
                load_calc_contract(path)
            except CalcContractError:
                continue
            except Exception as error:  # noqa: BLE001
                raise AssertionError(
                    f"{label}: raised {type(error).__name__} instead of CalcContractError"
                ) from error
        accepted.append(label)
    assert not accepted, "duplicate keys silently accepted at: " + ", ".join(accepted)


def test_a_duplicate_is_caught_even_where_the_two_values_are_identical() -> None:
    """Detection must not depend on the values differing."""
    _rejected_text(
        _duplicate_before('  required_visibility: "hidden"\n'),
        "two identical declarations of the same key",
    )


def test_the_strict_loader_neither_takes_first_nor_last() -> None:
    """Both orderings are refused, so no implicit resolution rule exists.

    Ordering matters to the hazard: last-wins means the WRONG value is the one a
    human reads, first-wins means it is the one the validator checks. Neither is
    an acceptable answer, so both arrangements must fail.
    """
    anchor = '        units: "SAR per unit"\n'
    original = _raw_contract()
    assert anchor in original

    for order in (
        ('        units: "USD"\n', anchor),          # wrong value read first
        (anchor, '        units: "USD"\n'),          # wrong value parsed last
    ):
        text = original.replace(anchor, "".join(order), 1)
        assert text.count('        units: "USD"\n') == 1
        _rejected_text(text, f"duplicate units in order {order}")


def test_the_strict_loader_is_still_a_safe_loader() -> None:
    """No arbitrary Python object construction, no unsafe tags."""
    text = _raw_contract().replace(
        'calc_contract_version: "1.0.0"',
        'calc_contract_version: !!python/object/apply:os.system ["echo unsafe"]',
    )
    with tempfile.TemporaryDirectory(prefix="pccm-unsafe-") as tmp:
        path = Path(tmp) / "unsafe.yaml"
        path.write_text(text, encoding="utf-8")
        try:
            load_calc_contract(path)
        except CalcContractError:
            return
    raise AssertionError("an unsafe YAML tag was constructed")


def test_ordinary_yaml_semantics_are_unchanged_by_the_strict_loader() -> None:
    """Scalars, lists and nested mappings still parse exactly as before."""
    calc = load_calc_contract(CALC_PATH)
    assert calc.phase4_cells == ("C10", "C11")                       # list of scalars
    assert calc.derived_status_labels[0] == "NOT CALCULATED"          # list of strings
    assert calc.tolerances.profiling_sum_absolute == 1e-9             # float scalar
    assert calc.tolerances.fx_rate_strictly_positive is True          # bool scalar
    assert calc.calc_state.field_by_key("last_successful_stamp").initial is None  # null
    assert calc.tolerances.conditioning_terms["i1"] == ("abs_a", "abs_b", "abs_c")


def test_malformed_yaml_is_reported_as_a_contract_error() -> None:
    with tempfile.TemporaryDirectory(prefix="pccm-badyaml-") as tmp:
        path = Path(tmp) / "broken.yaml"
        path.write_text("sheet:\n  name: '_Calc\n", encoding="utf-8")
        try:
            load_calc_contract(path)
        except CalcContractError:
            return
    raise AssertionError("malformed YAML was silently accepted")


def test_the_strict_parser_does_not_replace_the_schema_validators() -> None:
    """It is an ADDITIONAL boundary guard. Every earlier check must still fire."""
    checks: list[tuple[str, Callable[[dict[str, Any]], None]]] = [
        ("version lock", lambda d: d.__setitem__("calc_contract_version", "2.0.0")),
        ("semantic table key lock",
         lambda d: d["tables"].__setitem__("foo", d["tables"].pop("calc_fx"))),
        ("unknown key rejection", lambda d: d["sheet"].__setitem__("foo", "bar")),
        ("table schema lock",
         lambda d: _column(d, "tblCalcFX", "fx_to_sar").__setitem__("header", "Rate")),
        ("calc_state lock",
         lambda d: _state_field(d, "last_successful_stamp").__setitem__("value_type", "text")),
        ("calc_totals lock",
         lambda d: _totals_field(d, "a_nom").__setitem__("number_format", "0")),
        ("exact tolerances",
         lambda d: d["tolerances"].__setitem__("identity_absolute_floor", 1e-3)),
        ("conditioning terms",
         lambda d: d["tolerances"]["conditioning_terms"].__setitem__("i1", ["abs_d", "abs_e"])),
        ("authority-reference set",
         lambda d: d.__setitem__(
             "authority_references",
             [r for r in d["authority_references"] if r["concept"] != "FX convention"],
         )),
        ("hash-mathematics exclusion",
         lambda d: d["fingerprint"].__setitem__("note", "modulus 2147483647")),
        ("Phase-4 reservation",
         lambda d: d["phase4_reservation"].__setitem__("last_row", 9)),
        ("status axis lock",
         lambda d: d["state_labels"]["derived_status"].append("REFUSED")),
    ]
    for label, mutate in checks:
        _rejected(mutate, f"{label} no longer fires")


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
    # The same six values as the loader's locked copy, so the two cannot drift.
    assert LOCKED_TOLERANCES == {
        "profiling_sum_absolute": 1e-9,
        "identity_absolute_floor": 1e-6,
        "identity_relative_coefficient": 1e-12,
        "conditioning_scale_floor": 1.0,
        "fx_rate_strictly_positive": True,
        "growth_factor_strictly_positive": True,
    }


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


def test_the_exact_conditioning_terms_of_each_identity() -> None:
    """Each identity's scale sums ITS OWN absolute magnitudes."""
    tolerances = load_calc_contract(CALC_PATH).tolerances
    expected = {
        "i1": ("abs_a", "abs_b", "abs_c"),
        "i2": ("abs_c", "abs_d", "abs_e"),
        "i3a": ("sum_abs_annual_base", "abs_c"),
        "i3b": ("sum_abs_annual_risk", "abs_d"),
        "i3c": ("sum_abs_annual_total", "abs_e"),
        "i4a": ("sum_abs_annual_base", "abs_c"),
        "i4b": ("sum_abs_annual_risk", "abs_d"),
        "i4c": ("sum_abs_annual_total", "abs_e"),
    }
    assert tolerances.conditioning_terms == expected
    assert LOCKED_CONDITIONING_TERMS == expected


def test_every_tolerance_constant_is_individually_locked() -> None:
    """A tolerance edit is a numerical-design change, not a tuning knob."""
    changes: dict[str, Any] = {
        "profiling_sum_absolute": 1e-3,
        "identity_absolute_floor": 1e-3,
        "identity_relative_coefficient": 1e-9,
        "conditioning_scale_floor": 10,
        "fx_rate_strictly_positive": False,
        "growth_factor_strictly_positive": False,
    }
    assert set(changes) == set(LOCKED_TOLERANCES)
    for name, value in changes.items():

        def mutate(data: dict[str, Any], k: str = name, v: Any = value) -> None:
            data["tolerances"][k] = v

        _rejected(mutate, f"tolerance {name} changed")


def test_every_identity_conditioning_term_set_is_individually_locked() -> None:
    for identity in LOCKED_CONDITIONING_TERMS:

        def mutate(data: dict[str, Any], k: str = identity) -> None:
            data["tolerances"]["conditioning_terms"][k] = ["abs_a", "abs_e"]

        _rejected(mutate, f"conditioning terms for {identity} replaced")


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
