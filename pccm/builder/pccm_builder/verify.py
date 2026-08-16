"""Structural verification of a generated PCCM workbook.

This checks the artifact against the manifest that produced it. It is used both
by the builder (as a post-build gate) and by the Phase 1 test suite.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

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


def verify_workbook(path: str | Path, spec: WorkbookSpec) -> VerificationResult:
    """Verify the workbook at *path* against *spec*."""
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
        result.check(
            "workbook defines no defined names",
            len(list(workbook.defined_names)) == 0,
            ", ".join(list(workbook.defined_names)),
        )
        result.check(
            "no worksheet declares an Excel table",
            not _tables(workbook),
            "; ".join(_tables(workbook)),
        )
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


def _tables(workbook) -> list[str]:
    found: list[str] = []
    for worksheet in workbook.worksheets:
        for name in getattr(worksheet, "tables", {}):
            found.append(f"{worksheet.title}!{name}")
    return found


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
