"""Load and validate the PCCM workbook manifest.

The manifest is the structural authority. This module fails loudly on any
specification error and never repairs one silently.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

VALID_VISIBILITY = ("visible", "hidden", "veryHidden")
VALID_BLOCK_TYPES = ("section", "note")
VALID_BODIES = ("contract", "drivers", "structure")
CODENAME_RE = re.compile(r"^sh[A-Z][A-Za-z0-9]*$")


class SpecError(Exception):
    """Raised when the manifest is structurally invalid."""


@dataclass(frozen=True)
class SheetSpec:
    name: str
    codename: str
    visibility: str
    role: str
    purpose: str
    show_gridlines: bool
    title: str
    subtitle: str | None
    body: str | None
    freeze_panes: str | None
    column_widths: dict[str, float]
    blocks: list[dict[str, Any]] = field(default_factory=list)

    @property
    def is_visible(self) -> bool:
        return self.visibility == "visible"


@dataclass(frozen=True)
class WorkbookSpec:
    manifest_version: str
    model: dict[str, Any]
    workbook: dict[str, Any]
    presentation: dict[str, Any]
    sheets: list[SheetSpec]
    source_path: Path
    phase6_shell: dict[str, Any] = field(default_factory=dict)

    @property
    def sheet_names(self) -> list[str]:
        return [s.name for s in self.sheets]

    @property
    def active_sheet(self) -> str:
        return self.workbook["active_sheet"]

    @property
    def stage_a_filename(self) -> str:
        return self.workbook["stage_a_filename"]

    @property
    def contract_sheets(self) -> list[str]:
        """Sheets whose body is generated from the input contract."""
        return [s.name for s in self.sheets if s.body == "contract"]

    @property
    def driver_sheets(self) -> list[str]:
        """Sheets whose body is generated from the driver contract."""
        return [s.name for s in self.sheets if s.body == "drivers"]

    @property
    def structure_sheets(self) -> list[str]:
        """Sheets whose body is generated from the structure contract."""
        return [s.name for s in self.sheets if s.body == "structure"]

    def sheet(self, name: str) -> SheetSpec:
        for s in self.sheets:
            if s.name == name:
                return s
        raise SpecError(f"no sheet named {name!r} in manifest")


def _require(mapping: dict[str, Any], key: str, where: str) -> Any:
    if key not in mapping:
        raise SpecError(f"{where}: missing required key {key!r}")
    return mapping[key]


def _require_str(mapping: dict[str, Any], key: str, where: str) -> str:
    value = _require(mapping, key, where)
    if not isinstance(value, str) or not value.strip():
        raise SpecError(f"{where}: {key!r} must be a non-empty string, got {value!r}")
    return value


_SHELL_FORMULA_KEYS = ("formula", "nominal", "pv", "confidence_level_formula",
                       "quantile_nominal", "quantile_pv", "contingency_nominal",
                       "contingency_pv")


def _parse_phase6_shell(raw: dict[str, Any], path: Path) -> dict[str, Any]:
    """The Phase-6 publication shell: WHERE every label and formula sits.

    Shape only here; WHICH fields must exist is `sim_contract.yaml`'s, and the
    simulation loader cross-validates the two. What this does enforce is that
    every formula is a formula and every row is a row, because a label silently
    written into a formula cell is not something a later reader can see.
    """
    shell = raw.get("phase6_shell")
    if shell is None:
        return {}
    if not isinstance(shell, dict):
        raise SpecError(f"{path}: phase6_shell must be a mapping")
    for section in ("sim_data", "results"):
        if section not in shell:
            raise SpecError(f"{path}: phase6_shell omits {section!r}")
        if not isinstance(shell[section], dict):
            raise SpecError(f"{path}: phase6_shell.{section} must be a mapping")

    def walk(node: Any, where: str, formulas: bool) -> None:
        if isinstance(node, dict):
            for key, value in node.items():
                # `nominal` and `pv` are formula cells on Results and header TEXT
                # on _SimData, so the rule is scoped to where formulas live.
                if (formulas and key in _SHELL_FORMULA_KEYS
                        and not where.endswith(".headers")
                        and not where.endswith(".labels")):
                    if not isinstance(value, str) or not value.startswith("="):
                        raise SpecError(
                            f"{where}.{key} must be a formula beginning with '='"
                        )
                if key == "row" or key.endswith("_row"):
                    if not isinstance(value, int) or isinstance(value, bool) or value < 1:
                        raise SpecError(f"{where}.{key} must be a positive row number")
                walk(value, f"{where}.{key}", formulas)
        elif isinstance(node, list):
            for index, item in enumerate(node):
                walk(item, f"{where}[{index}]", formulas)

    walk(shell["sim_data"], f"{path}: phase6_shell.sim_data", False)
    walk(shell["results"], f"{path}: phase6_shell.results", True)
    _check_results_layout(shell["results"], path)
    if "dashboard" in shell:
        walk(shell["dashboard"], f"{path}: phase6_shell.dashboard", True)
        _check_dashboard_layout(shell["dashboard"], shell["results"], path)
    return shell


def _check_dashboard_layout(dashboard: dict[str, Any], results: dict[str, Any],
                            path: Path) -> None:
    """P8-2. The Dashboard sections must not overlap each other, must not reach
    the reserved chart region, and must mirror rows Results actually publishes.

    WHY EACH ONE IS HERE.

    OVERLAP is silent. Two sections whose rows intersect both render, the later
    one wins, and every check of the loser reads the winner's value - which is
    exactly the class of defect the Results annual/reconciliation collision
    check above was written for, one sheet along.

    THE CHART REGION is reserved for a step that has not happened. A summary
    that grew down into it would be overwritten the moment P8-3 draws there, and
    the failure would look like a chart bug rather than a layout one.

    A DANGLING SOURCE KEY is the P7-4 failure mode. The Dashboard names Results
    rows by block and key precisely so a moved row moves the mirror; a key that
    Results does not publish must fail the BUILD, naming the key, rather than
    reaching the renderer as a KeyError or - worse - resolving to nothing.
    """
    where = f"{path}: phase6_shell.dashboard"
    if dashboard.get("source_sheet") != results["sheet"]:
        raise SpecError(
            f"{where}.source_sheet is {dashboard.get('source_sheet')!r}; the Dashboard "
            f"mirrors {results['sheet']!r} and nothing else")
    template = str(dashboard.get("mirror_formula", ""))
    # THE BLANK GUARD IS NOT OPTIONAL. Excel reads an empty reference back as 0,
    # so a bare `=Results!$D$47` prints a fabricated zero under the words NOT
    # PRODUCED. The guard is the difference between an honest blank and a lie.
    if template.count("{ref}") != 2 or not template.startswith('=IF({ref}=""'):
        raise SpecError(
            f"{where}.mirror_formula must guard the reference against blank before "
            f"returning it; found {template!r}")

    available = {
        "run_stamp": {str(f["key"]) for f in results["run_stamp"]["fields"]},
        "summary": {str(m["key"]) for m in results["summary"]["metrics"]},
        "selected": {"confidence_level", "total", "contingency"},
        "state": {"distribution_state", "profile_state", "profile_px", "year_count"},
        "reconciliation": {str(e["key"]) for e in results["reconciliation"]["rows"]},
    }
    formats = dashboard["number_formats"]
    region = dashboard["chart_region"]
    reserved_top = int(region["heading_row"])
    occupied: dict[int, str] = {}

    for section in dashboard["sections"]:
        key = str(section["key"])
        rows = [int(section["row"]), int(section["note_row"])]
        if section.get("headers"):
            rows.append(int(section["header_row"]))
        first = int(section["first_row"])
        rows += [first + offset for offset in range(len(section["rows"]))]
        for row in rows:
            if row in occupied:
                raise SpecError(
                    f"{where}: sections {occupied[row]!r} and {key!r} both write row "
                    f"{row}")
            occupied[row] = key
            if row >= reserved_top:
                raise SpecError(
                    f"{where}: section {key!r} reaches row {row}, at or inside the "
                    f"chart region reserved from row {reserved_top}")
        for entry in section["rows"]:
            block, source = str(entry["source_block"]), str(entry["source_key"])
            if block not in available:
                raise SpecError(
                    f"{where}: section {key!r} mirrors block {block!r}, which Results "
                    f"does not publish")
            if source not in available[block]:
                raise SpecError(
                    f"{where}: section {key!r} mirrors {block}.{source!r}, which the "
                    f"Results {block} block does not publish")
            if str(entry["format"]) not in formats:
                raise SpecError(
                    f"{where}.number_formats: no format is declared for "
                    f"{entry['format']!r}")
    if int(region["note_row"]) <= reserved_top or int(region["last_row"]) <= int(
            region["note_row"]):
        raise SpecError(f"{where}.chart_region rows are not in ascending order")


def _check_results_layout(results: dict[str, Any], path: Path) -> None:
    """The Phase-8 sections must fit BELOW the accepted Phase-6 ones.

    THE ROWS ARE THE WHOLE CONTRACT BETWEEN THESE BLOCKS. Phase 6's run stamp,
    summary and selected-Px rows were accepted at fixed coordinates and a
    Windows harness reads them; a Phase-8 section that started one row too high
    would overwrite an accepted cell and every check of it would still pass,
    because both would be reading whatever landed there last.

    And a declared format must exist. A `format` key naming a number format the
    block does not carry is a KeyError deep inside the renderer, at which point
    the message is about a dictionary rather than about the manifest.
    """
    where = f"{path}: phase6_shell.results"
    annual = results.get("annual")
    if annual is None:
        return
    if "reconciliation" not in results or "number_formats" not in results:
        raise SpecError(
            f"{where}: the annual section needs both a reconciliation block and "
            "number_formats; a half-declared Phase-8 surface renders half a sheet")
    reconciliation = results["reconciliation"]
    formats = results["number_formats"]

    accepted_last = max(
        [int(f["row"]) for f in results["run_stamp"]["fields"]]
        + [int(m["row"]) for m in results["summary"]["metrics"]]
        + [int(results["selected"][key]) for key in
           ("confidence_level_row", "quantile_row", "contingency_row")]
    )
    if int(annual["heading_row"]) <= accepted_last:
        raise SpecError(
            f"{where}.annual: starts at row {annual['heading_row']}, which is inside "
            f"the accepted Phase-6 layout ending at row {accepted_last}")

    window_top = int(annual["first_row"])
    if int(annual["header_row"]) >= window_top:
        raise SpecError(
            f"{where}.annual: the header row must sit above the first record row")
    if int(reconciliation["heading_row"]) <= window_top:
        raise SpecError(
            f"{where}.reconciliation: starts at row {reconciliation['heading_row']}, "
            f"which is inside the annual record window beginning at {window_top}")

    used = [column["format"] for column in annual["columns"]]
    used += [row["format"] for row in reconciliation["rows"]]
    used += ["year"]
    missing = sorted({name for name in used if name not in formats})
    if missing:
        raise SpecError(f"{where}.number_formats: no format is declared for {missing}")
    for column in annual["columns"]:
        if column["source"] not in ("index", "profile"):
            raise SpecError(
                f"{where}.annual: column {column['key']!r} names source "
                f"{column['source']!r}, which is not an annual record block")


def load_spec(path: str | Path) -> WorkbookSpec:
    """Parse and fully validate the manifest at *path*."""
    path = Path(path)
    if not path.is_file():
        raise SpecError(f"manifest not found: {path}")

    with path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle)

    if not isinstance(raw, dict):
        raise SpecError(f"{path}: manifest root must be a mapping")

    manifest_version = _require_str(raw, "manifest_version", str(path))
    model = _require(raw, "model", str(path))
    workbook = _require(raw, "workbook", str(path))
    presentation = _require(raw, "presentation", str(path))
    raw_sheets = _require(raw, "sheets", str(path))
    shell = _parse_phase6_shell(raw, path)

    for key in ("name", "short_name", "model_version", "build_phase", "reporting_currency"):
        _require_str(model, key, f"{path}: model")
    for key in ("stage_a_filename", "active_sheet"):
        _require_str(workbook, key, f"{path}: workbook")
    for key in ("font_family", "sizes", "colors", "row_heights", "layout"):
        _require(presentation, key, f"{path}: presentation")

    if not isinstance(raw_sheets, list) or not raw_sheets:
        raise SpecError(f"{path}: 'sheets' must be a non-empty list")

    sheets = [_parse_sheet(entry, index, path) for index, entry in enumerate(raw_sheets)]

    _validate_unique(sheets, path)
    _validate_locked_order(sheets, workbook, path)
    _validate_active_sheet(sheets, workbook, path)

    return WorkbookSpec(
        manifest_version=manifest_version,
        model=model,
        workbook=workbook,
        presentation=presentation,
        sheets=sheets,
        source_path=path,
        phase6_shell=shell,
    )


def _parse_sheet(entry: Any, index: int, path: Path) -> SheetSpec:
    where = f"{path}: sheets[{index}]"
    if not isinstance(entry, dict):
        raise SpecError(f"{where}: each sheet must be a mapping")

    name = _require_str(entry, "name", where)
    where = f"{path}: sheet {name!r}"

    codename = _require_str(entry, "codename", where)
    if not CODENAME_RE.match(codename):
        raise SpecError(
            f"{where}: codename {codename!r} must match the sh<PascalCase> convention"
        )

    visibility = _require_str(entry, "visibility", where)
    if visibility not in VALID_VISIBILITY:
        raise SpecError(
            f"{where}: visibility {visibility!r} must be one of {VALID_VISIBILITY}"
        )

    show_gridlines = _require(entry, "show_gridlines", where)
    if not isinstance(show_gridlines, bool):
        raise SpecError(f"{where}: show_gridlines must be a boolean")

    widths = entry.get("column_widths") or {}
    if not isinstance(widths, dict):
        raise SpecError(f"{where}: column_widths must be a mapping of column letter to width")
    for column, width in widths.items():
        if not isinstance(column, str) or not column.isalpha():
            raise SpecError(f"{where}: column key {column!r} is not a column letter")
        if not isinstance(width, (int, float)) or width <= 0:
            raise SpecError(f"{where}: width for column {column} must be a positive number")

    body = entry.get("body")
    if body is not None and body not in VALID_BODIES:
        raise SpecError(f"{where}: body {body!r} must be omitted or one of {VALID_BODIES}")

    blocks = entry.get("blocks") or []
    if not isinstance(blocks, list):
        raise SpecError(f"{where}: blocks must be a list")
    if body is not None and blocks:
        raise SpecError(
            f"{where}: a {body}-bodied sheet must not also declare blocks; "
            "its contract is the single layout authority for that sheet"
        )
    if body is None and not blocks:
        raise SpecError(f"{where}: sheet has neither blocks nor a body contract")
    for position, block in enumerate(blocks):
        _validate_block(block, f"{where}: blocks[{position}]")

    return SheetSpec(
        name=name,
        codename=codename,
        visibility=visibility,
        role=_require_str(entry, "role", where),
        purpose=_require_str(entry, "purpose", where),
        show_gridlines=show_gridlines,
        title=_require_str(entry, "title", where),
        subtitle=entry.get("subtitle"),
        body=body,
        freeze_panes=entry.get("freeze_panes"),
        column_widths={str(k): float(v) for k, v in widths.items()},
        blocks=blocks,
    )


def _validate_block(block: Any, where: str) -> None:
    if not isinstance(block, dict):
        raise SpecError(f"{where}: block must be a mapping")
    block_type = _require_str(block, "type", where)
    if block_type not in VALID_BLOCK_TYPES:
        raise SpecError(f"{where}: type {block_type!r} must be one of {VALID_BLOCK_TYPES}")

    if block_type == "note":
        _require_str(block, "text", where)
        return

    _require_str(block, "title", where)
    rows = block.get("rows") or []
    if not isinstance(rows, list):
        raise SpecError(f"{where}: rows must be a list")
    for position, row in enumerate(rows):
        if not isinstance(row, dict):
            raise SpecError(f"{where}: rows[{position}] must be a mapping")
        _require_str(row, "label", f"{where}: rows[{position}]")

    items = block.get("list") or []
    if not isinstance(items, list):
        raise SpecError(f"{where}: list must be a list")
    for position, item in enumerate(items):
        if not isinstance(item, str) or not item.strip():
            raise SpecError(f"{where}: list[{position}] must be a non-empty string")


def _validate_unique(sheets: list[SheetSpec], path: Path) -> None:
    names = [s.name for s in sheets]
    duplicates = {n for n in names if names.count(n) > 1}
    if duplicates:
        raise SpecError(f"{path}: duplicate sheet names: {sorted(duplicates)}")

    codenames = [s.codename for s in sheets]
    duplicates = {c for c in codenames if codenames.count(c) > 1}
    if duplicates:
        raise SpecError(f"{path}: duplicate intended CodeNames: {sorted(duplicates)}")


def _validate_locked_order(
    sheets: list[SheetSpec], workbook: dict[str, Any], path: Path
) -> None:
    locked = workbook.get("locked_sheet_order")
    if locked is None:
        raise SpecError(f"{path}: workbook.locked_sheet_order is required")
    if not isinstance(locked, list) or not all(isinstance(n, str) for n in locked):
        raise SpecError(f"{path}: workbook.locked_sheet_order must be a list of strings")

    actual = [s.name for s in sheets]
    if actual != locked:
        raise SpecError(
            f"{path}: sheet order drifted from the architecture lock.\n"
            f"  locked:   {locked}\n"
            f"  manifest: {actual}"
        )


def _validate_active_sheet(
    sheets: list[SheetSpec], workbook: dict[str, Any], path: Path
) -> None:
    active = workbook["active_sheet"]
    match = next((s for s in sheets if s.name == active), None)
    if match is None:
        raise SpecError(f"{path}: active_sheet {active!r} is not one of the defined sheets")
    if not match.is_visible:
        raise SpecError(
            f"{path}: active_sheet {active!r} is {match.visibility!r}; "
            "a hidden sheet must never be the active sheet"
        )
