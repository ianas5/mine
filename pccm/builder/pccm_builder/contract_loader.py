"""Load and validate the PCCM input contract.

The contract is the semantic authority for model inputs, list masters, tables,
defined names and validation. Like the structural manifest, it fails loudly and
never repairs an invalid specification silently.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

from openpyxl.utils import get_column_letter

CELL_RE = re.compile(r"^([A-Z]{1,3})([1-9][0-9]{0,6})$")
DEFINED_NAME_RE = re.compile(r"^(inp|lst)[A-Z][A-Za-z0-9]*$")
TABLE_NAME_RE = re.compile(r"^tbl[A-Z][A-Za-z0-9]*$")

VALID_TYPES = ("text", "integer", "decimal", "percentage")
VALID_VALIDATION_KINDS = ("list", "whole", "decimal")


class ContractError(Exception):
    """Raised when the input contract is invalid."""


@dataclass(frozen=True)
class InputSpec:
    key: str
    sheet: str
    label: str
    defined_name: str
    label_cell: str
    cell: str
    type: str
    required: bool
    editable: bool
    default: Any
    number_format: str
    validation: dict[str, Any] | None
    note: str | None


@dataclass(frozen=True)
class TableColumnSpec:
    header: str
    number_format: str
    validation: dict[str, Any] | None


@dataclass(frozen=True)
class TableSpec:
    key: str
    sheet: str
    table_name: str
    header_row: int
    first_column: str
    data_rows: int
    editable: bool
    columns: list[TableColumnSpec]
    seed_rows: list[list[Any]] = field(default_factory=list)
    defined_name: str | None = None
    section: str | None = None
    section_row: int | None = None
    note: str | None = None
    note_row: int | None = None

    @property
    def first_data_row(self) -> int:
        return self.header_row + 1

    @property
    def last_data_row(self) -> int:
        return self.header_row + self.data_rows

    @property
    def first_col_index(self) -> int:
        return _column_index(self.first_column)

    @property
    def last_column(self) -> str:
        return get_column_letter(self.first_col_index + len(self.columns) - 1)

    @property
    def ref(self) -> str:
        return f"{self.first_column}{self.header_row}:{self.last_column}{self.last_data_row}"

    def column_letter(self, index: int) -> str:
        return get_column_letter(self.first_col_index + index)

    def data_range(self, index: int = 0) -> str:
        letter = self.column_letter(index)
        return f"{letter}{self.first_data_row}:{letter}{self.last_data_row}"

    def absolute_data_range(self, index: int = 0) -> str:
        letter = self.column_letter(index)
        return f"'{self.sheet}'!${letter}${self.first_data_row}:${letter}${self.last_data_row}"


@dataclass(frozen=True)
class SectionSpec:
    title: str
    row: int
    inputs: list[str] = field(default_factory=list)
    note: str | None = None
    note_row: int | None = None
    convention_row: int | None = None
    table: str | None = None


@dataclass(frozen=True)
class InputContract:
    contract_version: str
    conventions: dict[str, Any]
    inputs: dict[str, InputSpec]
    setup_sheet: str
    setup_intro: dict[str, Any]
    setup_sections: list[SectionSpec]
    config_sheet: str
    config_intro: dict[str, Any]
    tables: dict[str, TableSpec]
    config_tables: list[TableSpec]
    source_path: Path

    @property
    def fx_convention(self) -> str:
        return self.conventions["fx_convention"]

    @property
    def contract_sheets(self) -> set[str]:
        return {self.setup_sheet, self.config_sheet}

    def inputs_for_sheet(self, sheet: str) -> list[InputSpec]:
        return [i for i in self.inputs.values() if i.sheet == sheet]

    def tables_for_sheet(self, sheet: str) -> list[TableSpec]:
        return [t for t in self.all_tables if t.sheet == sheet]

    @property
    def all_tables(self) -> list[TableSpec]:
        return list(self.tables.values()) + list(self.config_tables)

    @property
    def list_defined_names(self) -> dict[str, str]:
        return {t.defined_name: t.absolute_data_range() for t in self.config_tables if t.defined_name}

    @property
    def input_defined_names(self) -> dict[str, str]:
        return {
            spec.defined_name: f"'{spec.sheet}'!${_col(spec.cell)}${_row(spec.cell)}"
            for spec in self.inputs.values()
        }


# ---------------------------------------------------------------------------
# loading
# ---------------------------------------------------------------------------
def load_contract(path: str | Path) -> InputContract:
    path = Path(path)
    if not path.is_file():
        raise ContractError(f"input contract not found: {path}")

    with path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle)
    if not isinstance(raw, dict):
        raise ContractError(f"{path}: contract root must be a mapping")

    where = str(path)
    contract_version = _req_str(raw, "contract_version", where)
    conventions = _req(raw, "conventions", where)
    for key in ("input_prefix", "list_prefix", "table_prefix", "fx_convention"):
        _req_str(conventions, key, f"{where}: conventions")

    inputs = _parse_inputs(_req(raw, "inputs", where), path)
    tables = _parse_tables(_req(raw, "tables", where), path)
    config_tables = _parse_config_tables(_req(raw, "config_tables", where), path)

    setup_layout = _req(raw, "setup_layout", where)
    config_layout = _req(raw, "config_layout", where)
    setup_sheet = _req_str(setup_layout, "sheet", f"{where}: setup_layout")
    config_sheet = _req_str(config_layout, "sheet", f"{where}: config_layout")
    sections = _parse_sections(_req(setup_layout, "sections", where), path)

    contract = InputContract(
        contract_version=contract_version,
        conventions=conventions,
        inputs=inputs,
        setup_sheet=setup_sheet,
        setup_intro=_req(setup_layout, "intro", f"{where}: setup_layout"),
        setup_sections=sections,
        config_sheet=config_sheet,
        config_intro=_req(config_layout, "intro", f"{where}: config_layout"),
        tables=tables,
        config_tables=config_tables,
        source_path=path,
    )

    _validate_unique_names(contract, path)
    _validate_sections(contract, path)
    _validate_validation_sources(contract, path)
    _validate_no_cell_collisions(contract, path)
    _validate_seed_rows(contract, path)
    return contract


def _parse_inputs(raw: Any, path: Path) -> dict[str, InputSpec]:
    if not isinstance(raw, dict) or not raw:
        raise ContractError(f"{path}: 'inputs' must be a non-empty mapping")
    result: dict[str, InputSpec] = {}
    for key, entry in raw.items():
        where = f"{path}: input {key!r}"
        if not isinstance(entry, dict):
            raise ContractError(f"{where}: must be a mapping")

        defined_name = _req_str(entry, "defined_name", where)
        if not DEFINED_NAME_RE.match(defined_name):
            raise ContractError(
                f"{where}: defined_name {defined_name!r} must match inp<PascalCase>"
            )
        type_ = _req_str(entry, "type", where)
        if type_ not in VALID_TYPES:
            raise ContractError(f"{where}: type {type_!r} must be one of {VALID_TYPES}")

        for cell_key in ("label_cell", "cell"):
            value = _req_str(entry, cell_key, where)
            if not CELL_RE.match(value):
                raise ContractError(f"{where}: {cell_key} {value!r} is not a valid cell reference")

        for flag in ("required", "editable"):
            if not isinstance(entry.get(flag), bool):
                raise ContractError(f"{where}: {flag!r} must be a boolean")

        validation = entry.get("validation")
        if validation is not None:
            _validate_validation_block(validation, where)

        if entry.get("editable") is False and entry.get("default") is None:
            raise ContractError(
                f"{where}: a model-controlled input must declare its locked default value"
            )

        result[key] = InputSpec(
            key=key,
            sheet=_req_str(entry, "sheet", where),
            label=_req_str(entry, "label", where),
            defined_name=defined_name,
            label_cell=entry["label_cell"],
            cell=entry["cell"],
            type=type_,
            required=bool(entry["required"]),
            editable=bool(entry["editable"]),
            default=entry.get("default"),
            number_format=_req_str(entry, "number_format", where),
            validation=validation,
            note=entry.get("note"),
        )
    return result


def _parse_tables(raw: Any, path: Path) -> dict[str, TableSpec]:
    if not isinstance(raw, dict) or not raw:
        raise ContractError(f"{path}: 'tables' must be a non-empty mapping")
    result: dict[str, TableSpec] = {}
    for key, entry in raw.items():
        where = f"{path}: table {key!r}"
        columns = _parse_columns(entry, where)
        result[key] = TableSpec(
            key=key,
            sheet=_req_str(entry, "sheet", where),
            table_name=_table_name(entry, where),
            header_row=_positive_int(entry, "header_row", where),
            first_column=_column(entry, where),
            data_rows=_positive_int(entry, "data_rows", where),
            editable=_bool(entry, "editable", where),
            columns=columns,
            seed_rows=entry.get("seed_rows") or [],
        )
    return result


def _parse_config_tables(raw: Any, path: Path) -> list[TableSpec]:
    if not isinstance(raw, list) or not raw:
        raise ContractError(f"{path}: 'config_tables' must be a non-empty list")
    result: list[TableSpec] = []
    for index, entry in enumerate(raw):
        where = f"{path}: config_tables[{index}]"
        if not isinstance(entry, dict):
            raise ContractError(f"{where}: must be a mapping")
        key = _req_str(entry, "key", where)
        where = f"{path}: config table {key!r}"

        defined_name = _req_str(entry, "defined_name", where)
        if not DEFINED_NAME_RE.match(defined_name) or not defined_name.startswith("lst"):
            raise ContractError(
                f"{where}: defined_name {defined_name!r} must match lst<PascalCase>"
            )

        values = entry.get("values") or []
        if not isinstance(values, list):
            raise ContractError(f"{where}: values must be a list")
        data_rows = _positive_int(entry, "data_rows", where)
        if len(values) > data_rows:
            raise ContractError(
                f"{where}: {len(values)} seeded values exceed data_rows={data_rows}"
            )
        editable = _bool(entry, "editable", where)
        if not editable and not values:
            raise ContractError(f"{where}: a locked list must declare its constant values")
        if not editable and len(values) != data_rows:
            raise ContractError(
                f"{where}: a locked list must size data_rows to its value count "
                f"({len(values)} values, data_rows={data_rows})"
            )

        result.append(
            TableSpec(
                key=key,
                sheet="Config",
                table_name=_table_name(entry, where),
                header_row=_positive_int(entry, "header_row", where),
                first_column=_column(entry, where),
                data_rows=data_rows,
                editable=editable,
                columns=[
                    TableColumnSpec(
                        header=_req_str(entry, "column_header", where),
                        number_format=_req_str(entry, "number_format", where),
                        validation=None,
                    )
                ],
                seed_rows=[[value] for value in values],
                defined_name=defined_name,
                section=_req_str(entry, "section", where),
                section_row=_positive_int(entry, "section_row", where),
                note=entry.get("note"),
                note_row=entry.get("note_row"),
            )
        )
    return result


def _parse_columns(entry: dict[str, Any], where: str) -> list[TableColumnSpec]:
    raw_columns = _req(entry, "columns", where)
    if not isinstance(raw_columns, list) or not raw_columns:
        raise ContractError(f"{where}: columns must be a non-empty list")
    columns: list[TableColumnSpec] = []
    for position, column in enumerate(raw_columns):
        col_where = f"{where}: columns[{position}]"
        if not isinstance(column, dict):
            raise ContractError(f"{col_where}: must be a mapping")
        validation = column.get("validation")
        if validation is not None:
            _validate_validation_block(validation, col_where)
        columns.append(
            TableColumnSpec(
                header=_req_str(column, "header", col_where),
                number_format=_req_str(column, "number_format", col_where),
                validation=validation,
            )
        )
    headers = [c.header for c in columns]
    if len(headers) != len(set(headers)):
        raise ContractError(f"{where}: duplicate column headers {headers}")
    return columns


def _parse_sections(raw: Any, path: Path) -> list[SectionSpec]:
    if not isinstance(raw, list) or not raw:
        raise ContractError(f"{path}: setup_layout.sections must be a non-empty list")
    sections: list[SectionSpec] = []
    for index, entry in enumerate(raw):
        where = f"{path}: setup_layout.sections[{index}]"
        if not isinstance(entry, dict):
            raise ContractError(f"{where}: must be a mapping")
        sections.append(
            SectionSpec(
                title=_req_str(entry, "title", where),
                row=_positive_int(entry, "row", where),
                inputs=entry.get("inputs") or [],
                note=entry.get("note"),
                note_row=entry.get("note_row"),
                convention_row=entry.get("convention_row"),
                table=entry.get("table"),
            )
        )
    return sections


# ---------------------------------------------------------------------------
# cross-cutting validation
# ---------------------------------------------------------------------------
def _validate_unique_names(contract: InputContract, path: Path) -> None:
    defined = [spec.defined_name for spec in contract.inputs.values()]
    defined += [t.defined_name for t in contract.config_tables if t.defined_name]
    duplicates = sorted({n for n in defined if defined.count(n) > 1})
    if duplicates:
        raise ContractError(f"{path}: duplicate defined names: {duplicates}")

    tables = [t.table_name for t in contract.all_tables]
    duplicates = sorted({n for n in tables if tables.count(n) > 1})
    if duplicates:
        raise ContractError(f"{path}: duplicate table names: {duplicates}")

    collisions = sorted(set(defined) & set(tables))
    if collisions:
        raise ContractError(f"{path}: names used for both a range and a table: {collisions}")


def _validate_sections(contract: InputContract, path: Path) -> None:
    seen: set[str] = set()
    for section in contract.setup_sections:
        for key in section.inputs:
            if key not in contract.inputs:
                raise ContractError(
                    f"{path}: section {section.title!r} references unknown input {key!r}"
                )
            if key in seen:
                raise ContractError(f"{path}: input {key!r} appears in more than one section")
            seen.add(key)
        if section.table is not None and section.table not in contract.tables:
            raise ContractError(
                f"{path}: section {section.title!r} references unknown table {section.table!r}"
            )
    setup_inputs = {k for k, v in contract.inputs.items() if v.sheet == contract.setup_sheet}
    missing = sorted(setup_inputs - seen)
    if missing:
        raise ContractError(f"{path}: Setup inputs not placed in any section: {missing}")


def _validate_validation_sources(contract: InputContract, path: Path) -> None:
    known = set(contract.list_defined_names)
    sources: list[tuple[str, str]] = []
    for spec in contract.inputs.values():
        if spec.validation and spec.validation.get("kind") == "list":
            sources.append((f"input {spec.key!r}", spec.validation["source"]))
    for table in contract.all_tables:
        for column in table.columns:
            if column.validation and column.validation.get("kind") == "list":
                sources.append(
                    (f"table {table.table_name!r} column {column.header!r}", column.validation["source"])
                )
    for owner, source in sources:
        if source not in known:
            raise ContractError(
                f"{path}: {owner} validates against {source!r}, which is not a declared "
                f"list defined name. Known: {sorted(known)}"
            )


def _validate_no_cell_collisions(contract: InputContract, path: Path) -> None:
    """No two contract-owned cells may occupy the same address on a sheet."""
    occupied: dict[str, dict[str, str]] = {}

    def claim(sheet: str, address: str, owner: str) -> None:
        sheet_cells = occupied.setdefault(sheet, {})
        if address in sheet_cells:
            raise ContractError(
                f"{path}: cell {sheet}!{address} claimed by both "
                f"{sheet_cells[address]} and {owner}"
            )
        sheet_cells[address] = owner

    for spec in contract.inputs.values():
        claim(spec.sheet, spec.label_cell, f"label of {spec.key!r}")
        claim(spec.sheet, spec.cell, f"value of {spec.key!r}")

    for section in contract.setup_sections:
        claim(contract.setup_sheet, f"B{section.row}", f"section {section.title!r}")
        if section.note_row:
            claim(contract.setup_sheet, f"B{section.note_row}", f"note of {section.title!r}")
        if section.convention_row:
            claim(contract.setup_sheet, f"B{section.convention_row}", f"convention of {section.title!r}")

    for table in contract.all_tables:
        if table.section_row:
            claim(table.sheet, f"B{table.section_row}", f"section of {table.table_name}")
        if table.note_row:
            claim(table.sheet, f"B{table.note_row}", f"note of {table.table_name}")
        for column_index in range(len(table.columns)):
            letter = table.column_letter(column_index)
            for row in range(table.header_row, table.last_data_row + 1):
                claim(table.sheet, f"{letter}{row}", f"table {table.table_name}")


def _validate_seed_rows(contract: InputContract, path: Path) -> None:
    for table in contract.all_tables:
        for index, row in enumerate(table.seed_rows):
            if not isinstance(row, list):
                raise ContractError(f"{path}: {table.table_name} seed_rows[{index}] must be a list")
            if len(row) != len(table.columns):
                raise ContractError(
                    f"{path}: {table.table_name} seed_rows[{index}] has {len(row)} values "
                    f"but the table has {len(table.columns)} columns"
                )
        if len(table.seed_rows) > table.data_rows:
            raise ContractError(
                f"{path}: {table.table_name} has more seed rows than data_rows"
            )


def _validate_validation_block(validation: Any, where: str) -> None:
    if not isinstance(validation, dict):
        raise ContractError(f"{where}: validation must be a mapping")
    kind = _req_str(validation, "kind", where)
    if kind not in VALID_VALIDATION_KINDS:
        raise ContractError(f"{where}: validation kind {kind!r} must be one of {VALID_VALIDATION_KINDS}")
    if kind == "list":
        _req_str(validation, "source", where)
    else:
        _req_str(validation, "operator", where)
        _req_str(validation, "formula1", where)


# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------
def _req(mapping: Any, key: str, where: str) -> Any:
    if not isinstance(mapping, dict) or key not in mapping:
        raise ContractError(f"{where}: missing required key {key!r}")
    return mapping[key]


def _req_str(mapping: Any, key: str, where: str) -> str:
    value = _req(mapping, key, where)
    if not isinstance(value, str) or not value.strip():
        raise ContractError(f"{where}: {key!r} must be a non-empty string, got {value!r}")
    return value


def _positive_int(mapping: dict[str, Any], key: str, where: str) -> int:
    value = _req(mapping, key, where)
    if not isinstance(value, int) or isinstance(value, bool) or value < 1:
        raise ContractError(f"{where}: {key!r} must be a positive integer, got {value!r}")
    return value


def _bool(mapping: dict[str, Any], key: str, where: str) -> bool:
    value = _req(mapping, key, where)
    if not isinstance(value, bool):
        raise ContractError(f"{where}: {key!r} must be a boolean, got {value!r}")
    return value


def _column(mapping: dict[str, Any], where: str) -> str:
    value = _req_str(mapping, "first_column", where)
    if not value.isalpha() or not value.isupper():
        raise ContractError(f"{where}: first_column {value!r} must be an upper-case column letter")
    return value


def _table_name(mapping: dict[str, Any], where: str) -> str:
    value = _req_str(mapping, "table_name", where)
    if not TABLE_NAME_RE.match(value):
        raise ContractError(f"{where}: table_name {value!r} must match tbl<PascalCase>")
    return value


def _col(cell: str) -> str:
    match = CELL_RE.match(cell)
    if match is None:
        raise ContractError(f"invalid cell reference {cell!r}")
    return match.group(1)


def _row(cell: str) -> int:
    match = CELL_RE.match(cell)
    if match is None:
        raise ContractError(f"invalid cell reference {cell!r}")
    return int(match.group(2))


def _column_index(letter: str) -> int:
    index = 0
    for char in letter:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index
