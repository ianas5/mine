"""Excel data validation.

Only validations justified by an already-locked rule are created. Data validation
is input infrastructure, not the Model Check engine: it guides entry at the point
of typing, and every advisory rule (duration > 25, iterations < 10000,
Base Year <= Start Year, required-ness) belongs to Model Check in a later phase.

Every rule permits a blank cell. A blank required input is a Model Check concern,
not something to block at the keyboard.

Validation is applied to USER-OWNED rows only. A table's locked seed rows carry
model invariants such as the SAR FX identity; the user does not own them, so
attaching user-input validation to them would misrepresent who controls the value.
"""

from __future__ import annotations

from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from .contract_loader import InputContract


def apply_validation(worksheets: dict[str, Worksheet], contract: InputContract) -> list[str]:
    """Attach every contract-declared validation. Returns a description per rule."""
    applied: list[str] = []

    for spec in contract.inputs.values():
        if not spec.validation:
            continue
        worksheet = worksheets[spec.sheet]
        dv = _build(spec.validation)
        worksheet.add_data_validation(dv)
        dv.add(spec.cell)          # address string; dv.add takes a coordinate
        applied.append(f"{spec.sheet}!{spec.cell} <- {_describe(spec.validation)}")

    for table in contract.all_tables:
        worksheet = worksheets[table.sheet]
        for index, column in enumerate(table.columns):
            if not column.validation:
                continue
            target = table.user_data_range(index)
            if target is None:
                # Wholly locked table, or every row is a locked identity row.
                continue
            dv = _build(column.validation)
            worksheet.add_data_validation(dv)
            dv.add(target)         # user-owned rows only, e.g. B29:B39
            applied.append(
                f"{table.sheet}!{target} ({table.table_name}.{column.header}) "
                f"<- {_describe(column.validation)}"
            )

    return applied


def _build(rule: dict) -> DataValidation:
    kind = rule["kind"]
    allow_blank = bool(rule.get("allow_blank", True))

    if kind == "list":
        dv = DataValidation(
            type="list",
            formula1=f"={rule['source']}",
            allow_blank=allow_blank,
        )
    else:
        dv = DataValidation(
            type=kind,
            operator=rule["operator"],
            formula1=rule["formula1"],
            formula2=rule.get("formula2"),
            allow_blank=allow_blank,
        )

    dv.showInputMessage = True
    dv.showErrorMessage = True
    dv.errorStyle = "stop"
    dv.promptTitle = rule.get("prompt_title")
    dv.prompt = rule.get("prompt")
    dv.errorTitle = rule.get("error_title")
    dv.error = rule.get("error")
    return dv


def _describe(rule: dict) -> str:
    if rule["kind"] == "list":
        return f"list from {rule['source']}"
    return f"{rule['kind']} {rule['operator']} {rule['formula1']}"
