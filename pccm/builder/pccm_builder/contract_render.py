"""Render the Setup and Config sheet bodies from the input contract.

Population only. Defined names live in ``names.py`` and data validation in
``validation.py``; neither is created here. No formula, calculation or business
rule is written by this module.
"""

from __future__ import annotations

from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from .contract_loader import InputContract, InputSpec, TableSpec
from .styling import StyleBook

NOTE_COLUMN = "E"


def render_setup(worksheet: Worksheet, contract: InputContract, styles: StyleBook) -> None:
    """Setup: scalar inputs grouped into sections, plus the FX rate table."""
    intro = contract.setup_intro
    _write(worksheet, f"B{intro['row']}", intro["text"], styles.note)

    for section in contract.setup_sections:
        _write(worksheet, f"B{section.row}", section.title, styles.section)
        worksheet.row_dimensions[section.row].height = styles.row_height("section")

        if section.note and section.note_row:
            _write(worksheet, f"B{section.note_row}", section.note, styles.note)
        if section.convention_row:
            _write(
                worksheet,
                f"B{section.convention_row}",
                f"Convention: {contract.fx_convention}",
                styles.note,
            )

        for key in section.inputs:
            _render_input(worksheet, contract.inputs[key], styles)

        if section.table:
            _render_table(worksheet, contract.tables[section.table], styles)


def render_config(worksheet: Worksheet, contract: InputContract, styles: StyleBook) -> None:
    """Config: one Excel Table per list master, each under its own section."""
    intro = contract.config_intro
    _write(worksheet, f"B{intro['row']}", intro["text"], styles.note)

    for table in contract.config_tables:
        if table.section and table.section_row:
            _write(worksheet, f"B{table.section_row}", table.section, styles.section)
            worksheet.row_dimensions[table.section_row].height = styles.row_height("section")
        if table.note and table.note_row:
            _write(worksheet, f"B{table.note_row}", table.note, styles.note)
        _render_table(worksheet, table, styles)


# ---------------------------------------------------------------------------
# building blocks
# ---------------------------------------------------------------------------
def _render_input(worksheet: Worksheet, spec: InputSpec, styles: StyleBook) -> None:
    _write(worksheet, spec.label_cell, spec.label, styles.label)

    cell = worksheet[spec.cell]
    cell.value = spec.default          # None leaves the cell genuinely blank
    cell.number_format = spec.number_format
    if spec.editable:
        styles.apply_input(cell)
    else:
        styles.apply_locked(cell)

    if spec.note:
        row = worksheet[spec.cell].row
        _write(worksheet, f"{NOTE_COLUMN}{row}", spec.note, styles.note)


def _render_table(worksheet: Worksheet, table: TableSpec, styles: StyleBook) -> None:
    for index, column in enumerate(table.columns):
        letter = table.column_letter(index)

        header = worksheet[f"{letter}{table.header_row}"]
        header.value = column.header
        styles.apply_table_header(header)

        for offset in range(table.data_rows):
            row = table.first_data_row + offset
            cell = worksheet[f"{letter}{row}"]
            if offset < len(table.seed_rows):
                cell.value = table.seed_rows[offset][index]
            cell.number_format = column.number_format
            # Locked rows are model invariants (e.g. the SAR identity) or a wholly
            # locked constant list. Either way the user does not own them.
            if table.is_locked_row(offset):
                styles.apply_locked(cell)
            else:
                styles.apply_input(cell)

    excel_table = Table(displayName=table.table_name, ref=table.ref)
    excel_table.tableStyleInfo = TableStyleInfo(
        name=styles.table_style_name,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
        showColumnStripes=False,
    )
    worksheet.add_table(excel_table)


def _write(worksheet: Worksheet, address: str, value, font) -> None:
    cell = worksheet[address]
    cell.value = value
    cell.font = font
