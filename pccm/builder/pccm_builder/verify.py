"""Structural verification of a generated PCCM workbook.

This checks the artifact against the manifest that produced it. It is used both
by the builder (as a post-build gate) and by the Phase 1 test suite.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from .contract_loader import InputContract
from .spec_loader import WorkbookSpec


@dataclass
class VerificationResult:
    passed: list[str] = field(default_factory=list)
    failures: list[str] = field(default_factory=list)

    def check(self, description: str, condition: bool, detail: str = "") -> bool:
        if condition:
            self.passed.append(description)
        else:
            self.failures.append(f"{description}{f' -- {detail}' if detail else ''}")
        return condition

    @property
    def ok(self) -> bool:
        return not self.failures

    def report(self) -> str:
        lines = [f"  [PASS] {item}" for item in self.passed]
        lines += [f"  [FAIL] {item}" for item in self.failures]
        lines.append("")
        lines.append(f"  {len(self.passed)} passed, {len(self.failures)} failed")
        return "\n".join(lines)


def verify_workbook(
    path: str | Path, spec: WorkbookSpec, contract: InputContract | None = None
) -> VerificationResult:
    """Verify the workbook at *path* against the manifest and the input contract."""
    path = Path(path)
    result = VerificationResult()

    if not result.check("artifact exists", path.is_file(), str(path)):
        return result

    result.check(
        "artifact is a genuine OOXML (ZIP) package",
        zipfile.is_zipfile(path),
    )
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
    result.check(
        "no vbaProject.bin is present in this Stage-A .xlsx",
        not any(n.lower().endswith("vbaproject.bin") for n in names),
        ", ".join(n for n in names if "vba" in n.lower()),
    )
    result.check(
        "no external link parts are present",
        not any("externalLink" in n for n in names),
        ", ".join(n for n in names if "externalLink" in n),
    )
    result.check(
        "no data connection parts are present",
        not any("connections" in n.lower() for n in names),
        ", ".join(n for n in names if "connections" in n.lower()),
    )

    workbook = load_workbook(path)
    try:
        expected = spec.sheet_names
        actual = workbook.sheetnames

        result.check(
            f"workbook contains exactly {len(expected)} worksheets",
            len(actual) == len(expected),
            f"found {len(actual)}",
        )
        result.check("sheet names and order match the manifest", actual == expected,
                     f"expected {expected}, found {actual}")
        result.check(
            "no unexpected worksheet exists",
            set(actual) <= set(expected),
            f"unexpected: {sorted(set(actual) - set(expected))}",
        )

        for sheet_spec in spec.sheets:
            if sheet_spec.name not in workbook.sheetnames:
                result.check(f"sheet {sheet_spec.name!r} exists", False)
                continue
            worksheet = workbook[sheet_spec.name]
            result.check(
                f"sheet {sheet_spec.name!r} visibility is {sheet_spec.visibility!r}",
                worksheet.sheet_state == sheet_spec.visibility,
                f"found {worksheet.sheet_state!r}",
            )
            result.check(
                f"sheet {sheet_spec.name!r} shows its title",
                _contains(worksheet, sheet_spec.title),
            )

        active = workbook.active
        result.check(
            f"active sheet is {spec.active_sheet!r}",
            active is not None and active.title == spec.active_sheet,
            f"found {active.title if active is not None else None!r}",
        )
        result.check(
            "the active sheet is visible",
            active is not None and active.sheet_state == "visible",
        )

        result.check(
            "no worksheet contains a formula",
            not _formula_cells(workbook),
            "; ".join(_formula_cells(workbook)[:5]),
        )
        expected_names = set(contract.input_defined_names) | set(contract.list_defined_names) \
            if contract is not None else set()
        found_names = set(workbook.defined_names)
        result.check(
            "only contract-declared defined names exist",
            found_names == expected_names,
            f"unexpected {sorted(found_names - expected_names)}, "
            f"missing {sorted(expected_names - found_names)}",
        )
        expected_tables = (
            {t.table_name for t in contract.all_tables} if contract is not None else set()
        )
        found_tables = {name for _, name in _tables(workbook)}
        result.check(
            "only contract-declared Excel Tables exist",
            found_tables == expected_tables,
            f"expected {sorted(expected_tables)}, found {sorted(found_tables)}",
        )

        if contract is not None:
            _verify_contract(result, workbook, contract)
    finally:
        workbook.close()

    return result


def _contains(worksheet, text: str) -> bool:
    for row in worksheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value == text:
                return True
    return False


def _formula_cells(workbook) -> list[str]:
    found: list[str] = []
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    found.append(f"{worksheet.title}!{cell.coordinate}")
    return found


def _tables(workbook) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    for worksheet in workbook.worksheets:
        for name in getattr(worksheet, "tables", {}):
            found.append((worksheet.title, name))
    return found


def _verify_contract(result: VerificationResult, workbook, contract: InputContract) -> None:
    """Every contract promise must be observable in the built artifact."""
    for spec in contract.inputs.values():
        worksheet = workbook[spec.sheet]
        result.check(
            f"input {spec.key!r} label at {spec.sheet}!{spec.label_cell}",
            worksheet[spec.label_cell].value == spec.label,
            f"found {worksheet[spec.label_cell].value!r}",
        )
        result.check(
            f"input {spec.key!r} default at {spec.sheet}!{spec.cell}",
            worksheet[spec.cell].value == spec.default,
            f"expected {spec.default!r}, found {worksheet[spec.cell].value!r}",
        )
        reference = contract.input_defined_names[spec.defined_name]
        result.check(
            f"defined name {spec.defined_name} -> {spec.sheet}!{spec.cell}",
            workbook.defined_names[spec.defined_name].attr_text == reference,
            f"found {workbook.defined_names[spec.defined_name].attr_text!r}",
        )

    for table in contract.all_tables:
        worksheet = workbook[table.sheet]
        tables = getattr(worksheet, "tables", {})
        if not result.check(
            f"table {table.table_name} exists on {table.sheet}", table.table_name in tables
        ):
            continue
        result.check(
            f"table {table.table_name} ref is {table.ref}",
            tables[table.table_name].ref == table.ref,
            f"found {tables[table.table_name].ref}",
        )
        for index, column in enumerate(table.columns):
            address = f"{table.column_letter(index)}{table.header_row}"
            result.check(
                f"table {table.table_name} header {column.header!r} at {address}",
                worksheet[address].value == column.header,
                f"found {worksheet[address].value!r}",
            )
        for offset, seed in enumerate(table.seed_rows):
            for index, value in enumerate(seed):
                address = f"{table.column_letter(index)}{table.first_data_row + offset}"
                result.check(
                    f"table {table.table_name} seed {address} = {value!r}",
                    worksheet[address].value == value,
                    f"found {worksheet[address].value!r}",
                )
        if table.defined_name:
            result.check(
                f"defined name {table.defined_name} covers {table.table_name} data body",
                workbook.defined_names[table.defined_name].attr_text
                == table.absolute_data_range(),
                f"found {workbook.defined_names[table.defined_name].attr_text!r}",
            )

    # Model invariants: the declared identity values must be exactly right.
    for identity in contract.model_invariants["locked_identities"]:
        table = contract.table_by_name(identity["table"])
        worksheet = workbook[table.sheet]
        row = table.first_data_row + identity["row"] - 1
        for index, value in enumerate(identity["values"]):
            address = f"{table.column_letter(index)}{row}"
            result.check(
                f"model invariant {table.table_name} {address} = {value!r}",
                worksheet[address].value == value,
                f"found {worksheet[address].value!r}",
            )

    validations = {
        f"{ws.title}!{dv.sqref}": dv
        for ws in workbook.worksheets
        for dv in ws.data_validations.dataValidation
    }
    # No validation may target a locked identity row.
    for table in contract.all_tables:
        if not table.locked_seed_rows:
            continue
        worksheet = workbook[table.sheet]
        locked_rows = range(table.first_data_row, table.first_user_row)
        targeted = {
            str(cell)
            for dv in worksheet.data_validations.dataValidation
            for rng in dv.sqref.ranges
            for cell in rng.cells
        }
        offenders = [
            f"{table.column_letter(i)}{row}"
            for row in locked_rows
            for i in range(len(table.columns))
            if (table.column_letter(i), row) in {
                (c.split("$")[0] if "$" in c else "".join(ch for ch in c if ch.isalpha()),
                 int("".join(ch for ch in c if ch.isdigit())))
                for c in targeted
            }
        ]
        result.check(
            f"no data validation targets {table.table_name} locked identity rows",
            not offenders,
            ", ".join(offenders),
        )
    result.check(
        "data validation rules were created",
        len(validations) > 0,
        f"found {len(validations)}",
    )


def structural_digest(path: str | Path) -> str:
    """A normalised, order-sensitive description of workbook structure.

    Deliberately excludes ZIP metadata and file timestamps, so two builds of the
    same source compare equal without requiring byte-identical archives.
    """
    workbook = load_workbook(Path(path))
    try:
        parts: list[str] = []
        for worksheet in workbook.worksheets:
            parts.append(
                f"SHEET|{worksheet.title}|{worksheet.sheet_state}|"
                f"grid={worksheet.sheet_view.showGridLines}|freeze={worksheet.freeze_panes}"
            )
            for column, dimension in sorted(worksheet.column_dimensions.items()):
                parts.append(f"COL|{worksheet.title}|{column}|{dimension.width}")
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        parts.append(f"CELL|{worksheet.title}|{cell.coordinate}|{cell.value!r}")
        active = workbook.active
        parts.append(f"ACTIVE|{active.title if active is not None else None}")
        return "\n".join(parts)
    finally:
        workbook.close()
