"""Render the Cost Lines and Risk Register driver tables.

Population only: headers, reserved blank rows, per-column ownership styling and
column widths. Data validation is applied in ``validation.py``.

Nothing here writes a value into a data cell. In particular no identifier is
allocated: the ID columns are left genuinely blank because permanent-ID lifecycle
belongs to Stage B, and faking it with a row number or a formula would be worse
than leaving it empty.
"""

from __future__ import annotations

from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from .driver_loader import DriverRegister
from .styling import StyleBook


def render_register(
    worksheet: Worksheet, register: DriverRegister, styles: StyleBook
) -> None:
    if register.intro and register.intro_row:
        _write(worksheet, f"B{register.intro_row}", register.intro, styles.note)
    _write(worksheet, f"B{register.section_row}", register.section, styles.section)
    worksheet.row_dimensions[register.section_row].height = styles.row_height("section")
    if register.note and register.note_row:
        _write(worksheet, f"B{register.note_row}", register.note, styles.note)

    for index, column in enumerate(register.columns):
        letter = register.column_letter(index)
        worksheet.column_dimensions[letter].width = column.width

        header = worksheet[f"{letter}{register.header_row}"]
        header.value = column.header
        styles.apply_table_header(header)

        for row in range(register.first_data_row, register.last_data_row + 1):
            cell = worksheet[f"{letter}{row}"]
            cell.number_format = column.number_format
            # Column-level ownership: identity columns are model-controlled.
            if column.editable:
                styles.apply_input(cell)
            else:
                styles.apply_locked(cell)

    excel_table = Table(displayName=register.table_name, ref=register.ref)
    excel_table.tableStyleInfo = TableStyleInfo(
        name=styles.table_style_name,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
        showColumnStripes=False,
    )
    worksheet.add_table(excel_table)

    _apply_conditional_formatting(worksheet, register, styles)


def _apply_conditional_formatting(
    worksheet: Worksheet, register: DriverRegister, styles: StyleBook
) -> None:
    """Presentation only.

    Greys the Most Likely cell when the row's Distribution is Uniform, because
    Uniform has no Most Likely parameter. This constrains nothing: the cell still
    accepts input and carries no data validation. Behavioural disabling is a
    Stage B UI concern.
    """
    for rule in register.conditional_formatting:
        target = register.letter_of(rule.target_column)
        when = register.letter_of(rule.when_column)
        first, last = register.first_data_row, register.last_data_row
        formula = f'${when}{first}="{rule.equals}"'
        worksheet.conditional_formatting.add(
            f"{target}{first}:{target}{last}",
            FormulaRule(formula=[formula], fill=styles.not_applicable_fill, stopIfTrue=False),
        )


def _write(worksheet: Worksheet, address: str, value, font) -> None:
    cell = worksheet[address]
    cell.value = value
    cell.font = font
