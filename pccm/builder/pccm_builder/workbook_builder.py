"""Create the PCCM Stage A workbook from the manifest and the input contract.

Responsibilities are deliberately split across modules:
    workbook_builder  orchestration: sheets, visibility, active sheet, metadata
    contract_render   Setup and Config bodies (inputs, tables)
    names             defined names
    validation        data validation
    styling           presentation tokens
    spec_loader       structural manifest
    contract_loader   input contract
    verify            structural verification

No business rule or calculation belongs in any of them.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal
from typing import Any

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .contract_loader import InputContract
from .contract_render import render_config, render_setup
from .driver_loader import DriverContract, validate_against_input_contract
from .driver_render import render_register
from .names import apply_defined_names
from .spec_loader import SheetSpec, WorkbookSpec
from .structure_loader import StructureContract, validate_structure_against
from .calc_loader import CalcContract
# The Phase-8 adapters wrap the Phase-7 accessors under this prefix; the
# structural contract declares the four resulting procedure names.
RESULTS_STATE_PREFIX = "PCCM_Results"
from .calc_render import render_calc_workspace
from .structure_render import render_applied_timeline, render_grid, render_identity
from .styling import StyleBook
from .validation import apply_validation

BUILDER_VERSION = "0.5.0"
DEFAULT_SHEET_TITLE = "Sheet"
TIMESTAMP_ENV_VAR = "PCCM_BUILD_TIMESTAMP"


@dataclass(frozen=True)
class BuildMetadata:
    """Build-time provenance.

    None of these values is a computational input. They record how the artifact
    was produced and are kept distinct from:
      * model_version   - the version of the model design itself
      * builder_version - the version of this build tooling
      * run metadata    - which belongs to a Monte Carlo run, not to a build
    """

    model_version: str
    build_phase: str
    builder_version: str
    build_timestamp: str
    manifest_version: str
    contract_version: str
    driver_contract_version: str
    structure_contract_version: str

    @classmethod
    def create(
        cls,
        spec: WorkbookSpec,
        contract: InputContract,
        drivers: DriverContract,
        structure: StructureContract,
    ) -> "BuildMetadata":
        return cls(
            model_version=spec.model["model_version"],
            build_phase=spec.model["build_phase"],
            builder_version=BUILDER_VERSION,
            build_timestamp=resolve_build_timestamp(),
            manifest_version=spec.manifest_version,
            contract_version=contract.contract_version,
            driver_contract_version=drivers.version,
            structure_contract_version=structure.version,
        )

    def as_rows(self) -> list[tuple[str, str]]:
        return [
            ("PCCM Model Version", self.model_version),
            ("Build Phase", self.build_phase),
            ("Builder Version", self.builder_version),
            ("Build Timestamp (UTC)", self.build_timestamp),
            ("Source Manifest Version", self.manifest_version),
            ("Input Contract Version", self.contract_version),
            ("Driver Contract Version", self.driver_contract_version),
            ("Structure Contract Version", self.structure_contract_version),
        ]


def resolve_build_timestamp() -> str:
    """UTC build timestamp, overridable for reproducible builds.

    Setting PCCM_BUILD_TIMESTAMP makes two builds of the same source produce
    structurally identical workbooks, which is what the reproducibility test
    relies on.
    """
    override = os.environ.get(TIMESTAMP_ENV_VAR)
    if override:
        return override
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")


def build_workbook(
    spec: WorkbookSpec,
    contract: InputContract,
    drivers: DriverContract,
    structure: StructureContract,
    calc: CalcContract | None = None,
    sim: Any = None,
) -> tuple[Workbook, BuildMetadata]:
    """Create the Stage A workbook from the manifest and every contract.

    `calc` is the Phase-5 calculation contract. It is optional ONLY so that
    isolated Phase 1-4 unit tests can still build the structural workbook they
    were written against; `build_stage_a.py` always supplies it, and the
    production Stage-A path therefore always renders and verifies the Phase-5
    workspace. Passing `None` builds a workbook with no `_Calc` Phase-5 blocks,
    which post-build verification with a contract will reject.
    """
    _assert_consistent(spec, contract, drivers, structure)

    styles = StyleBook(spec.presentation)
    metadata = BuildMetadata.create(spec, contract, drivers, structure)

    workbook = Workbook()
    default_sheet = workbook.active

    for sheet_spec in spec.sheets:
        worksheet = workbook.create_sheet(title=sheet_spec.name)
        _apply_presentation(worksheet, sheet_spec, styles)
        _write_header(worksheet, sheet_spec, styles)
        if sheet_spec.body == "contract":
            if sheet_spec.name == contract.setup_sheet:
                render_setup(worksheet, contract, styles)
                # Setup carries two authorities: the input contract owns the entered
                # inputs and the FX table, the structure contract appends the applied
                # timeline below them. The loaders prove the two areas cannot overlap.
                render_applied_timeline(worksheet, structure, styles)
            else:
                render_config(worksheet, contract, styles)
        elif sheet_spec.body == "drivers":
            render_register(worksheet, drivers.register_for_sheet(sheet_spec.name), styles)
        elif sheet_spec.body == "structure":
            grid = structure.grid_for_sheet(sheet_spec.name)
            if grid is not None:
                render_grid(worksheet, grid, structure, styles)
            else:
                render_identity(worksheet, structure, styles)
                # `_Calc` carries two authorities: Phase 4 owns rows 1-11 above,
                # and the calculation contract owns everything from row 13 down.
                # The calc loader proves the two areas cannot intersect.
                if calc is not None and sheet_spec.name == calc.sheet:
                    render_calc_workspace(worksheet, calc, styles)
        else:
            _populate_blocks(worksheet, sheet_spec, styles, metadata)

        # THE PHASE-6 PUBLICATION SHELL. Labels, headers and presentation
        # formulas only - no iteration row, no snapshot, no digest. It is
        # materialised here rather than by the run precisely so a successful
        # `_SimData` commit can never be followed by a failed Results write.
        if sim is not None and spec.phase6_shell:
            render_phase6_shell(worksheet, sheet_spec, spec.phase6_shell, sim, contract,
                                styles, calc, structure)

    # Remove openpyxl's default sheet only after the real sheets exist, so the
    # workbook is never momentarily empty.
    if default_sheet is not None and default_sheet.title == DEFAULT_SHEET_TITLE:
        workbook.remove(default_sheet)
    if DEFAULT_SHEET_TITLE in workbook.sheetnames:
        raise RuntimeError("the default openpyxl sheet was not removed cleanly")

    # Visibility is applied after creation so the active sheet can be set while
    # every sheet is still visible.
    workbook.active = workbook.sheetnames.index(spec.active_sheet)
    for sheet_spec in spec.sheets:
        workbook[sheet_spec.name].sheet_state = sheet_spec.visibility

    apply_defined_names(workbook, contract, structure)
    apply_validation(
        {name: workbook[name] for name in workbook.sheetnames}, contract, drivers
    )

    _apply_document_properties(workbook, spec, metadata)
    return workbook, metadata


def _assert_consistent(
    spec: WorkbookSpec,
    contract: InputContract,
    drivers: DriverContract | None = None,
    structure: StructureContract | None = None,
) -> None:
    """The specifications must agree before anything is rendered."""
    # Cross-spec: the reporting currency is declared in both files. They must not
    # be allowed to drift apart silently.
    manifest_currency = spec.model["reporting_currency"]
    if manifest_currency != contract.reporting_currency:
        raise RuntimeError(
            "reporting currency disagrees between specifications: "
            f"workbook.yaml model.reporting_currency={manifest_currency!r}, "
            f"input_contract.yaml model_invariants.reporting_currency="
            f"{contract.reporting_currency!r}"
        )

    declared = set(spec.contract_sheets)
    used = contract.contract_sheets
    if declared != used:
        raise RuntimeError(
            "manifest and input contract disagree about contract-bodied sheets: "
            f"manifest says {sorted(declared)}, contract targets {sorted(used)}"
        )
    known = set(spec.sheet_names)
    for table in contract.all_tables:
        if table.sheet not in known:
            raise RuntimeError(
                f"input contract table {table.table_name!r} targets unknown sheet {table.sheet!r}"
            )
    for input_spec in contract.inputs.values():
        if input_spec.sheet not in known:
            raise RuntimeError(
                f"input {input_spec.key!r} targets unknown sheet {input_spec.sheet!r}"
            )

    if drivers is None:
        return

    declared_drivers = set(spec.driver_sheets)
    if declared_drivers != drivers.sheets:
        raise RuntimeError(
            "manifest and driver contract disagree about driver-bodied sheets: "
            f"manifest says {sorted(declared_drivers)}, "
            f"driver contract targets {sorted(drivers.sheets)}"
        )
    for register in drivers.all_registers:
        if register.sheet not in known:
            raise RuntimeError(
                f"driver register {register.table_name!r} targets unknown sheet "
                f"{register.sheet!r}"
            )
    validate_against_input_contract(drivers, contract)

    if structure is None:
        return

    declared_structure = set(spec.structure_sheets)
    if declared_structure != structure.owned_sheets:
        raise RuntimeError(
            "manifest and structure contract disagree about structure-bodied sheets: "
            f"manifest says {sorted(declared_structure)}, "
            f"structure contract owns {sorted(structure.owned_sheets)}"
        )
    for sheet in sorted(structure.owned_sheets | {structure.setup_sheet}):
        if sheet not in known:
            raise RuntimeError(f"structure contract targets unknown sheet {sheet!r}")
    if structure.setup_sheet != contract.setup_sheet:
        raise RuntimeError(
            "manifest and structure contract disagree about the Setup sheet: the applied "
            f"timeline targets {structure.setup_sheet!r}, the input contract owns "
            f"{contract.setup_sheet!r}"
        )
    for button in structure.buttons:
        if button.sheet not in known:
            raise RuntimeError(
                f"button {button.shape_name!r} targets unknown sheet {button.sheet!r}"
            )
    validate_structure_against(structure, contract, drivers)


def _apply_document_properties(
    workbook: Workbook, spec: WorkbookSpec, metadata: BuildMetadata
) -> None:
    properties = workbook.properties
    properties.title = spec.model["name"]
    properties.creator = f"{spec.model['short_name']} builder {metadata.builder_version}"
    properties.lastModifiedBy = properties.creator
    properties.category = metadata.build_phase
    properties.description = (
        f"{spec.model['short_name']} model version {metadata.model_version}; "
        f"manifest {metadata.manifest_version}; built {metadata.build_timestamp}."
    )
    # Deterministic document timestamps keep repeated builds comparable.
    fixed = datetime(2000, 1, 1, tzinfo=timezone.utc).replace(tzinfo=None)
    properties.created = fixed
    properties.modified = fixed


def _apply_presentation(
    worksheet: Worksheet, sheet_spec: SheetSpec, styles: StyleBook
) -> None:
    worksheet.sheet_view.showGridLines = sheet_spec.show_gridlines
    for column, width in sheet_spec.column_widths.items():
        worksheet.column_dimensions[column].width = width
    if sheet_spec.freeze_panes:
        worksheet.freeze_panes = sheet_spec.freeze_panes


def _write_header(worksheet: Worksheet, sheet_spec: SheetSpec, styles: StyleBook) -> None:
    """Title, subtitle and the thin rule. Common to every sheet."""
    layout = styles.layout
    label_col = layout.label_column

    _write(worksheet, f"{label_col}{layout.title_row}", sheet_spec.title, styles.title)
    worksheet.row_dimensions[layout.title_row].height = styles.row_height("title")

    if sheet_spec.subtitle:
        _write(
            worksheet,
            f"{label_col}{layout.subtitle_row}",
            sheet_spec.subtitle,
            styles.subtitle,
        )
        worksheet.row_dimensions[layout.subtitle_row].height = styles.row_height("subtitle")

    # A thin rule under the header, drawn across the used column span.
    for column in _rule_columns(sheet_spec):
        worksheet[f"{column}{layout.rule_row}"].border = styles.rule
    worksheet.row_dimensions[layout.rule_row].height = styles.row_height("spacer")


def render_phase6_shell(
    worksheet: Worksheet,
    sheet_spec: SheetSpec,
    shell: dict[str, Any],
    sim: Any,
    contract: InputContract,
    styles: StyleBook,
    calc: CalcContract | None = None,
    structure: StructureContract | None = None,
) -> None:
    """Materialise the empty publication shell.

    Every coordinate comes from `sim_contract.yaml`; this function reads the
    manifest only for WHERE things sit on Results and for the wording. Nothing
    written here is simulation output: no iteration record, no snapshot value,
    no fingerprint and no digest.
    """
    raw = sim.raw
    if sheet_spec.name == raw["sim_data"]["sheet"]:
        _render_sim_data_shell(worksheet, shell["sim_data"], raw, sim, contract, styles)
    elif sheet_spec.name == "Results":
        _render_results_shell(worksheet, shell["results"], styles, raw, calc, structure)
    elif sheet_spec.name == shell["sensitivity"]["sheet"]:
        _render_sensitivity_shell(worksheet, shell["sensitivity"], raw, styles)
    elif "dashboard" in shell and sheet_spec.name == shell["dashboard"]["sheet"]:
        _render_dashboard_shell(worksheet, shell["dashboard"], shell["results"], styles)


# ===========================================================================
# PHASE 8, STEP 2 - THE DASHBOARD EXECUTIVE SUMMARY
# ===========================================================================
# ONE GEOMETRY AUTHORITY, AND IT IS THE RESULTS BLOCK. Every cell this renderer
# writes is `=IF(Results!$D$nn="","",Results!$D$nn)`, and it never learns `nn`
# from the dashboard block: the manifest names a Results BLOCK and KEY, and the
# resolver below turns that pair into a row by reading the same `results:`
# structure that built the Results sheet minutes earlier. A Results row that
# moves takes this sheet with it; a key that stops existing fails the build by
# name rather than silently mirroring the wrong cell, which is the P7-4 failure
# mode - a hand-written address that went stale in silence and reported "Not
# produced for this run" forever.
#
# AND THE MIRROR IS THE WHOLE IMPLEMENTATION. There is no arithmetic here: no
# subtraction of a base from a total, no sum of a profile, no comparison to an
# allowance, no state word and no verdict. Whatever Results says, this sheet
# says, including the blank - `IF(ref="","",ref)` is what stops Excel reading an
# empty Results cell back as a hard 0 and printing a fabricated zero beside the
# words NOT PRODUCED.


def _dashboard_row_index(results: dict[str, Any]) -> dict[str, dict[str, int]]:
    """Every addressable Results row, by block and key.

    Built from the results block itself so the two can never disagree. The
    reconciliation rows are positional on Results - first_row plus the index of
    the entry - and are resolved here exactly the way the Results renderer
    resolves them, rather than being re-counted from a different starting row.
    """
    annual = results["annual"]
    reconciliation = results["reconciliation"]
    selected = results["selected"]
    return {
        "run_stamp": {str(field["key"]): int(field["row"])
                      for field in results["run_stamp"]["fields"]},
        "summary": {str(metric["key"]): int(metric["row"])
                    for metric in results["summary"]["metrics"]},
        "selected": {
            "confidence_level": int(selected["confidence_level_row"]),
            # THE TWO LADDERS, NAMED APART HERE TOO. `total` is the summary
            # block's rung at the selected level; `contingency` is that same
            # money less the deterministic base. A dashboard that resolved both
            # to one row would present one under the other's name and nothing on
            # the sheet would contradict it.
            "total": int(selected["quantile_row"]),
            "contingency": int(selected["contingency_row"]),
        },
        "state": {key: int(annual[f"{key}_row"])
                  for key in ("distribution_state", "profile_state",
                              "profile_px", "year_count")},
        "reconciliation": {str(entry["key"]): int(reconciliation["first_row"]) + index
                           for index, entry in enumerate(reconciliation["rows"])},
    }


def _dashboard_labels(results: dict[str, Any]) -> dict[str, dict[str, str]]:
    """The label Results already shows for each row, so this sheet never types
    a second one. Which rung `quantile_10` spells is the simulation contract's
    to say; repeating "P90" here would be a second declaration that a contract
    move would falsify in silence."""
    reconciliation = results["reconciliation"]
    return {
        "run_stamp": {str(f["key"]): str(f["label"])
                      for f in results["run_stamp"]["fields"]},
        "summary": {str(m["key"]): str(m["label"])
                    for m in results["summary"]["metrics"]},
        # KEYED THE WAY THE ROW INDEX KEYS THEM. Results names its own label
        # `quantile`; the row index calls that row `total`, because that is what
        # the money in it is. One spelling, resolved here, so a section entry
        # cannot address a row under one name and a label under another.
        "selected": {
            "confidence_level": str(results["selected"]["labels"]["confidence_level"]),
            "total": str(results["selected"]["labels"]["quantile"]),
            "contingency": str(results["selected"]["labels"]["contingency"]),
        },
        "state": {str(k): str(v) for k, v in results["annual"]["labels"].items()},
        "reconciliation": {str(e["key"]): str(e["label"])
                           for e in reconciliation["rows"]},
    }


def _render_dashboard_shell(
    worksheet: Worksheet, block: dict[str, Any], results: dict[str, Any],
    styles: StyleBook,
) -> None:
    label_col = block["label_column"]
    nominal_col = block["nominal_column"]
    pv_col = block["pv_column"]
    source_sheet = block["source_sheet"]
    template = block["mirror_formula"]
    formats = block["number_formats"]
    rows_by_block = _dashboard_row_index(results)
    labels_by_block = _dashboard_labels(results)

    source_columns = {
        "nominal": results["nominal_column"],
        "pv": results["pv_column"],
    }

    def mirror(source_block: str, source_key: str, measure: str) -> str:
        try:
            row = rows_by_block[source_block][source_key]
        except KeyError as error:
            raise ValueError(
                f"the Dashboard mirrors {source_block}.{source_key}, which the "
                f"Results block does not publish"
            ) from error
        reference = f"{source_sheet}!${source_columns[measure]}${row}"
        return template.format(ref=reference)

    for section in block["sections"]:
        _write(worksheet, f"{label_col}{section['row']}", section["title"], styles.section)
        worksheet.row_dimensions[int(section["row"])].height = styles.row_height("section")
        _write(worksheet, f"{label_col}{section['note_row']}", section["note"], styles.note)

        if section.get("headers"):
            header_row = int(section["header_row"])
            for column, key in ((label_col, "label"), (nominal_col, "nominal"),
                                (pv_col, "pv")):
                cell = worksheet[f"{column}{header_row}"]
                cell.value = section["headers"][key]
                styles.apply_table_header(cell)

        row = int(section["first_row"])
        for entry in section["rows"]:
            source_block = str(entry["source_block"])
            source_key = str(entry["source_key"])
            label = labels_by_block[source_block][source_key]
            _write(worksheet, f"{label_col}{row}", label, styles.label)
            # THE HEADLINE PAIR IS THE ONLY THING BOLDED, and it is bolded
            # rather than recomputed. Emphasis is presentation; the number is
            # still the same mirror as every other row on the sheet.
            font = styles.value_locked if entry.get("emphasis") else styles.value
            measures = ("nominal", "pv") if section.get("headers") else ("nominal",)
            for measure in measures:
                column = nominal_col if measure == "nominal" else pv_col
                cell = worksheet[f"{column}{row}"]
                cell.value = mirror(source_block, source_key, measure)
                cell.font = font
                cell.number_format = formats[str(entry["format"])]
            row += 1

    # THE RESERVED REGION, AND IT STAYS EMPTY. A heading and a note; no chart
    # object, no series, no anchor and no placeholder cell. P8-2 draws nothing.
    region = block["chart_region"]
    _write(worksheet, f"{label_col}{region['heading_row']}", region["heading"],
           styles.section)
    worksheet.row_dimensions[int(region["heading_row"])].height = styles.row_height("section")
    _write(worksheet, f"{label_col}{region['note_row']}", region["note"], styles.note)


def _render_sensitivity_shell(
    worksheet: Worksheet, block: dict[str, Any], raw: dict[str, Any],
    styles: StyleBook,
) -> None:
    """The Ranked Drivers table, as LOOKUPS into the persisted block.

    Not one cell computes anything. Every row reads the `_SimData` sensitivity
    records of the ACTIVE bank at a fixed offset, and blanks itself beyond the
    persisted record count - which is what stops a previous, longer result from
    leaving surplus rows on display when a later model has fewer drivers.

    The coordinates come from `sim_contract.yaml`; this function knows only
    where on the sheet they go.
    """
    records = raw["sim_data"]["sensitivity_records"]
    banks = records["banks"]
    stamp = records["stamp"]
    first_record_row = int(records["first_record_row"])
    count_row = next(f["row"] for f in stamp["fields"] if f["key"] == "record_count")
    published_row = next(f["row"] for f in stamp["fields"] if f["key"] == "published")
    label_col = block["label_column"]
    sheet = raw["sim_data"]["sheet"]
    active = f"{sheet}!$D$30"

    _write(worksheet, f"{label_col}{block['heading_row']}", block["heading"], styles.section)
    _write(worksheet, f"{label_col}{block['note_row']}", block["note"], styles.note)
    _write(worksheet, f"{label_col}{block['availability_row']}",
           block["availability_label"], styles.label)
    _write(worksheet, f"{block['columns'][1]['column']}{block['availability_row']}",
           block["availability_formula"], styles.value)

    for column in block["columns"]:
        _write(worksheet, f"{column['column']}{block['header_row']}",
               column["header"], styles.label)

    def cell(bank: str, offset: int, row_index: int) -> str:
        letter = _shift_column(banks[bank]["first_column"], offset)
        return f"{sheet}!${letter}${first_record_row + row_index}"

    def stamp_cell(bank: str, row: int) -> str:
        return f"{sheet}!${stamp['bank_value_columns'][bank]}${row}"

    for row_index in range(int(block["row_window"])):
        sheet_row = int(block["first_row"]) + row_index
        for column in block["columns"]:
            offset = int(column["offset"])
            # BLANK BEYOND THE COUNT. The persisted count is what bounds the
            # authoritative result; the window is only how much of it can show.
            formula = (
                f'=IF({active}="","",'
                f'IF(IF({active}="A",{stamp_cell("A", published_row)},'
                f'{stamp_cell("B", published_row)})<>"PUBLISHED","",'
                f'IF({row_index + 1}>IF({active}="A",{stamp_cell("A", count_row)},'
                f'{stamp_cell("B", count_row)}),"",'
                f'IF({active}="A",{cell("A", offset, row_index)},'
                f'{cell("B", offset, row_index)}))))'
            )
            _write(worksheet, f"{column['column']}{sheet_row}", formula, styles.value)


def _shift_column(letter: str, offset: int) -> str:
    value = 0
    for char in letter.upper():
        value = value * 26 + (ord(char) - ord("A") + 1)
    value += offset
    out = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        out = chr(ord("A") + remainder) + out
    return out


def _render_sim_data_shell(
    worksheet: Worksheet, block: dict[str, Any], raw: dict[str, Any], sim: Any,
    contract: InputContract, styles: StyleBook,
) -> None:
    identity = raw["sim_data"]["run_identity"]
    label_col = identity["label_column"]
    shared_col = identity["value_column"]
    banks = identity["bank_value_columns"]

    _write(worksheet, f"{label_col}{block['heading_row']}", block["identity_heading"],
           styles.section)
    _write(worksheet, f"{label_col}{block['note_row']}", block["identity_note"], styles.note)

    for field in identity["fields"]:
        _write(worksheet, f"{label_col}{field['row']}", field["label"], styles.label)
        if field["group"] == "snapshot":
            continue
        if field.get("initial") is not None:
            _write(worksheet, f"{shared_col}{field['row']}", field["initial"], styles.value)

    # The persisted summary and contingency blocks: labels and bank headers only.
    ladder = _shell_ladder(sim, contract)
    summary = raw["sim_data"]["summary_statistics"]
    _render_persisted_block(worksheet, summary, summary["metrics"], ladder,
                            block["summary_heading"], block["summary_headers"],
                            block["heading_row"], block["note_row"], styles)
    contingency = raw["sim_data"]["contingency_ladder"]
    _render_persisted_block(worksheet, contingency, contingency["rungs"], ladder,
                            block["contingency_heading"], block["contingency_headers"],
                            block["heading_row"], block["note_row"], styles)

    records = raw["sim_data"]["iteration_records"]
    _write(worksheet, f"{label_col}{block['iteration_heading_row']}",
           block["iteration_heading"], styles.section)
    _write(worksheet, f"{label_col}{block['iteration_note_row']}",
           block["iteration_note"], styles.note)
    for bank, columns in records["banks"].items():
        for key, column in columns.items():
            _write(worksheet, f"{column}{records['header_row']}",
                   block["bank_headers"][bank][key], styles.label)

    for column, width in (block.get("column_widths") or {}).items():
        worksheet.column_dimensions[column].width = width


def _shell_ladder(sim: Any, contract: InputContract) -> tuple[str, ...]:
    """The ladder, resolved from its OWNER through the accepted authority."""
    from .sim_oracle import resolve_percentile_ladder

    return tuple(resolve_percentile_ladder(sim, contract).ordered)


def _render_persisted_block(
    worksheet: Worksheet, block: dict[str, Any], entries: list[dict[str, Any]],
    ladder: tuple[str, ...], heading: str, headers: dict[str, Any],
    heading_row: int, note_row: int, styles: StyleBook,
) -> None:
    label_col = block["label_column"]
    _write(worksheet, f"{label_col}{heading_row}", heading, styles.section)
    for bank, columns in block["bank_value_columns"].items():
        for measure, column in columns.items():
            _write(worksheet, f"{column}{note_row}", headers[bank][measure], styles.label)
    for entry in entries:
        label = entry.get("label")
        if label is None:
            # THE LADDER LABEL COMES FROM ITS OWNER, never from this contract.
            label = ladder[int(entry["key"].split("_")[1]) - 1]
        _write(worksheet, f"{label_col}{entry['row']}", label, styles.label)


def _render_results_shell(
    worksheet: Worksheet, block: dict[str, Any], styles: StyleBook,
    raw: dict[str, Any] | None = None, calc: CalcContract | None = None,
    structure: StructureContract | None = None,
) -> None:
    label_col = block["label_column"]
    nominal_col = block["nominal_column"]
    pv_col = block["pv_column"]

    for section in block["sections"]:
        _write(worksheet, f"{label_col}{section['row']}", section["title"], styles.section)
        _write(worksheet, f"{label_col}{section['note_row']}", section["note"], styles.note)

    for field in block["run_stamp"]["fields"]:
        _write(worksheet, f"{label_col}{field['row']}", field["label"], styles.label)
        _write(worksheet, f"{nominal_col}{field['row']}", field["formula"], styles.value)

    summary = block["summary"]
    _write(worksheet, f"{label_col}{summary['header_row']}", summary["headers"]["label"],
           styles.label)
    _write(worksheet, f"{nominal_col}{summary['header_row']}", summary["headers"]["nominal"],
           styles.label)
    _write(worksheet, f"{pv_col}{summary['header_row']}", summary["headers"]["pv"],
           styles.label)
    money = block["number_formats"]["money"]
    for metric in summary["metrics"]:
        _write(worksheet, f"{label_col}{metric['row']}", metric["label"], styles.label)
        _write(worksheet, f"{nominal_col}{metric['row']}", metric["nominal"], styles.value)
        _write(worksheet, f"{pv_col}{metric['row']}", metric["pv"], styles.value)
        worksheet[f"{nominal_col}{metric['row']}"].number_format = money
        worksheet[f"{pv_col}{metric['row']}"].number_format = money

    selected = block["selected"]
    _write(worksheet, f"{label_col}{selected['confidence_level_row']}",
           selected["labels"]["confidence_level"], styles.label)
    _write(worksheet, f"{nominal_col}{selected['confidence_level_row']}",
           selected["confidence_level_formula"], styles.value)
    _write(worksheet, f"{label_col}{selected['quantile_row']}",
           selected["labels"]["quantile"], styles.label)
    _write(worksheet, f"{nominal_col}{selected['quantile_row']}",
           selected["quantile_nominal"], styles.value)
    _write(worksheet, f"{pv_col}{selected['quantile_row']}",
           selected["quantile_pv"], styles.value)
    _write(worksheet, f"{label_col}{selected['contingency_row']}",
           selected["labels"]["contingency"], styles.label)
    _write(worksheet, f"{nominal_col}{selected['contingency_row']}",
           selected["contingency_nominal"], styles.value)
    _write(worksheet, f"{pv_col}{selected['contingency_row']}",
           selected["contingency_pv"], styles.value)
    # THE TWO MONEY ROWS OF THE SELECTED BLOCK. The confidence level above them
    # is a LABEL - P80 is not a quantity - and is deliberately left alone.
    for row in (selected["quantile_row"], selected["contingency_row"]):
        worksheet[f"{nominal_col}{row}"].number_format = money
        worksheet[f"{pv_col}{row}"].number_format = money

    if raw is not None and "annual" in block:
        window = _annual_row_window(structure)
        # THE WINDOW IS SIZED BY A CONTRACT THIS FILE DOES NOT OWN. If the
        # structural year maximum grows, the annual table grows with it - and
        # the reconciliation below has a FIXED heading row. Nothing else would
        # notice the collision: the later write simply wins and both sections
        # would still "render". So it is refused here, in the build, naming the
        # two numbers that disagree.
        last = int(block["annual"]["first_row"]) + window - 1
        heading = int(block["reconciliation"]["heading_row"])
        if last >= heading:
            raise ValueError(
                f"the annual record window reaches row {last} but the reconciliation "
                f"heading is at row {heading}; the structural year maximum grew past "
                "the space the Results layout reserves for it")
        _render_annual_section(worksheet, block, styles, raw, window)
        _render_reconciliation_section(worksheet, block, styles, raw, calc, window)


# ===========================================================================
# PHASE 8, STEP 1 - THE ANNUAL CASH FLOW AND THE RECONCILIATION
# ===========================================================================
# PRESENTATION, AND NOTHING ELSE. Every cell below is a LOOKUP into the annual
# result Phase 7 published, or ordinary display arithmetic over cells that are
# already on this sheet. Nothing here computes a percentile, replays a run,
# blends a profile, resolves a selector or constructs a contingency: those have
# owners, and a worksheet that repeated any of them would be a second engine
# that no test of the VBA would ever see.
#
# AND NOT ONE ADDRESS IS TYPED. The Sensitivity availability line writes its
# `_SimData` addresses as literals, and P7-4 is what that cost: the persisted
# block moved and the formula did not, so a sheet reported "Not produced for
# this run" after a run that had just succeeded. Every column letter, stamp row
# and state word below is resolved from `sim_contract.yaml` at build time, and
# the two identity tolerances from `calc_contract.yaml`.


class _AnnualAddresses:
    """Where the published annual answer lives, resolved from the contract.

    One object so the four formula builders cannot disagree about which bank
    selector, which stamp row or which record column they are reading.
    """

    def __init__(self, raw: dict[str, Any]) -> None:
        self.sheet = str(raw["sim_data"]["sheet"])
        identity = raw["sim_data"]["run_identity"]
        self._identity_rows = {f["key"]: int(f["row"]) for f in identity["fields"]}
        self._identity_banks = identity["bank_value_columns"]
        self._shared_column = str(identity["value_column"])
        annual = raw["sim_data"]["annual_records"]
        self._annual = annual
        self._stamp_banks = annual["stamp"]["bank_value_columns"]
        self._stamp_rows = {f["key"]: int(f["row"]) for f in annual["stamp"]["fields"]}
        self.first_record_row = int(annual["first_record_row"])
        handoff = annual["handoff"]
        # THE ONLY TWO STATE WORDS THE SHEET STILL NEEDS, and neither is used to
        # DECIDE anything. `not_produced` blanks the record window over a verdict
        # the accessor reached; `profile_current` decides whether the
        # reconciliation status needs its "not the current answer" qualifier.
        # Every other state word left this file with the decision tree.
        self.not_produced = str(handoff["distribution_states"][0])
        self.profile_current = str(handoff["profile_states"][1])
        self._accessors = [str(entry["name"]) for entry in handoff["accessors"]]

    # The published bank decides every read below, and it is a cell.
    def active_bank(self) -> str:
        return (f'{self.sheet}!${self._shared_column}$'
                f'{self._identity_rows["active_bank"]}')

    def shared(self, key: str) -> str:
        """A row the two banks share - the derived status, the counters."""
        return f'{self.sheet}!${self._shared_column}${self._identity_rows[key]}'

    def _switch(self, banks: dict[str, str], row: int) -> str:
        return (f'IF({self.active_bank()}="A",'
                f'{self.sheet}!${banks["A"]}${row},{self.sheet}!${banks["B"]}${row})')

    def run(self, key: str) -> str:
        """A field of the published run's own identity."""
        return self._switch(self._identity_banks, self._identity_rows[key])

    def stamp(self, key: str) -> str:
        """A field of the annual publication stamp."""
        return self._switch(self._stamp_banks, self._stamp_rows[key])

    def state_procedures(self) -> dict[str, str]:
        """The Phase-8 adapter for each Phase-7 handoff accessor, by position.

        The contract declares the accessors in one order - distributions,
        profile, Px, year count - and the adapter for each is that name with the
        presentation prefix. Deriving the four rather than listing them keeps a
        renamed accessor from silently leaving a stale wrapper on the sheet.
        """
        keys = ("distribution_state", "profile_state", "profile_px", "year_count")
        if len(self._accessors) != len(keys):
            raise ValueError(
                f"the handoff declares {len(self._accessors)} accessors; the Results "
                f"state block presents {len(keys)}")
        return {key: RESULTS_STATE_PREFIX + name[len("PCCM_"):]
                for key, name in zip(keys, self._accessors)}

    def record(self, source: str, field: str, offset: int) -> str:
        """One cell of one annual record, `offset` rows into the block."""
        columns = (self._annual["index_columns"] if source == "index"
                   else self._annual["selected_px_profile_columns"])
        banks = {bank: columns[bank][field] for bank in ("A", "B")}
        return self._switch(banks, self.first_record_row + offset)


def _annual_row_window(structure: StructureContract | None) -> int:
    """How many record rows the sheet must be able to show.

    THE STRUCTURAL MAXIMUM ON GENERATED PROJECT-YEAR COLUMNS, which is what
    bounds the annual block - never a duration this project happens to have run.
    A four-year model and a two-hundred-year model use the same sheet, and the
    stamped year count blanks everything past the answer.
    """
    if structure is None:
        raise ValueError(
            "the Results annual table needs the structural contract: its row window "
            "is the maximum generated project-year count and must not be written here")
    window = int(structure.limits.max_generated_year_columns)
    if window < 1:
        raise ValueError(f"the structural year maximum is {window}, which shows nothing")
    return window


def _annual_state_formulas(at: _AnnualAddresses, block: dict[str, Any],
                           value_col: str) -> dict[str, str]:
    """The four state lines above the table - ASKED, NOT DERIVED.

    THE CORRECTION THIS BLOCK EXISTS FOR. These four cells used to rebuild the
    annual state out of `_SimData`: compare the publication marker, compare five
    stamp fields against the published run, read the persisted simulation
    status, then compare the stamped Px against the selector. It was a
    presentation layer owning half a semantic, and it was wrong in two specific
    ways. It read the PERSISTED simulation status, so after an ordinary model
    change the banner kept saying CURRENT until some later operation happened to
    re-evaluate it. And it implemented the selector arm of the profile rule while
    the accepted accessor also weighs whether the stamp agrees with itself - so a
    stamp whose label and probability disagreed would have been shown as current.

    Now each cell calls the Phase-7 accessor and prints what it says. There is no
    marker comparison, no identity test, no selector arm and no state word in any
    formula this function builds. What the sheet still does with the answer -
    blanking the record window when nothing was produced - is presentation over a
    verdict somebody else reached, which is the line this block is meant to sit on.
    """
    del block, value_col  # the layout no longer feeds the state formulas
    return {key: f"={procedure}()" for key, procedure in at.state_procedures().items()}


def _render_annual_section(
    worksheet: Worksheet, block: dict[str, Any], styles: StyleBook,
    raw: dict[str, Any], window: int,
) -> None:
    annual = block["annual"]
    label_col = block["label_column"]
    value_col = block["nominal_column"]
    formats = block["number_formats"]
    at = _AnnualAddresses(raw)

    _write(worksheet, f"{label_col}{annual['heading_row']}", annual["heading"], styles.section)
    worksheet.row_dimensions[int(annual["heading_row"])].height = styles.row_height("section")
    _write(worksheet, f"{label_col}{annual['note_row']}", annual["note"], styles.note)

    formulas = _annual_state_formulas(at, block, value_col)
    for key in ("distribution_state", "profile_state", "profile_px", "year_count"):
        row = int(annual[f"{key}_row"])
        _write(worksheet, f"{label_col}{row}", annual["labels"][key], styles.label)
        cell = worksheet[f"{value_col}{row}"]
        cell.value = formulas[key]
        cell.font = styles.value
        if key == "year_count":
            cell.number_format = formats["year"]

    header_row = int(annual["header_row"])
    for column in annual["columns"]:
        cell = worksheet[f"{column['column']}{header_row}"]
        cell.value = column["header"]
        styles.apply_table_header(cell)

    # THE WINDOW IS THE STRUCTURAL MAXIMUM, not a duration anybody has run. A
    # row beyond the stamped year count blanks itself, so a longer previous
    # answer can never leave surplus years on display.
    distribution_cell = f"${value_col}${annual['distribution_state_row']}"
    year_count_cell = f"${value_col}${annual['year_count_row']}"
    first_row = int(annual["first_row"])
    for offset in range(window):
        sheet_row = first_row + offset
        for column in annual["columns"]:
            lookup = at.record(column["source"], column["field"], offset)
            cell = worksheet[f"{column['column']}{sheet_row}"]
            cell.value = (f'=IF(OR({distribution_cell}="{at.not_produced}",'
                          f'{offset + 1}>{year_count_cell}),"",{lookup})')
            cell.font = styles.value
            cell.number_format = formats[column["format"]]


def _render_reconciliation_section(
    worksheet: Worksheet, block: dict[str, Any], styles: StyleBook,
    raw: dict[str, Any], calc: CalcContract | None, window: int,
) -> None:
    """`sum_y Profile_Px(y) = Total Px`, checked on the stored values.

    THE TOTAL IS READ FROM THE TOTAL, and this is where W5 went wrong once: the
    summary block's quantile rungs are the percentile of the iteration TOTALS,
    while the contingency block's rungs are `selected_px_total - deterministic
    base` - a smaller number by exactly the base. The row this reads is the
    Selected Px row, never the Contingency row below it.

    THE VERDICT IS TAKEN UNROUNDED. The cells are formatted for a reader; the
    comparison is between the stored Doubles, against the project's own identity
    allowance, and no profile is scaled to make it pass.
    """
    reconciliation = block["reconciliation"]
    annual = block["annual"]
    label_col = block["label_column"]
    nominal_col = block["nominal_column"]
    pv_col = block["pv_column"]
    formats = block["number_formats"]
    at = _AnnualAddresses(raw)
    if calc is None:
        raise ValueError(
            "the Results reconciliation needs the calculation contract: its identity "
            "allowance is calc_contract.yaml's and must not be written here")

    _write(worksheet, f"{label_col}{reconciliation['heading_row']}",
           reconciliation["heading"], styles.section)
    worksheet.row_dimensions[int(reconciliation["heading_row"])].height = (
        styles.row_height("section"))
    _write(worksheet, f"{label_col}{reconciliation['note_row']}",
           reconciliation["note"], styles.note)

    header_row = int(reconciliation["header_row"])
    for column, key in ((label_col, "label"), (nominal_col, "nominal"), (pv_col, "pv")):
        cell = worksheet[f"{column}{header_row}"]
        cell.value = reconciliation["headers"][key]
        styles.apply_table_header(cell)

    rows = {entry["key"]: int(reconciliation["first_row"]) + index
            for index, entry in enumerate(reconciliation["rows"])}
    selected = block["selected"]
    distribution_cell = f"${nominal_col}${annual['distribution_state_row']}"
    profile_state_cell = f"${nominal_col}${annual['profile_state_row']}"
    first = int(annual["first_row"])
    last = first + window - 1
    profile_columns = {column["field"]: column["column"]
                       for column in annual["columns"] if column["source"] == "profile"}

    tolerances = calc.tolerances
    absolute = _excel_number(tolerances.identity_absolute_floor)
    relative = _excel_number(tolerances.identity_relative_coefficient)
    scale_floor = _excel_number(tolerances.conditioning_scale_floor)
    verdicts = reconciliation["verdicts"]

    for measure, display_col in (("nominal", nominal_col), ("pv", pv_col)):
        profile_range = (f'${profile_columns[measure]}${first}'
                         f':${profile_columns[measure]}${last}')
        total = f'${display_col}${rows["total"]}'
        profile_sum = f'${display_col}${rows["profile_sum"]}'
        difference = f'${display_col}${rows["difference"]}'
        allowance = f'${display_col}${rows["allowance"]}'
        blank_guard = f'OR({total}="",{profile_sum}="")'
        # SUMIF IGNORES THE BLANKED ROWS. ABS over the range would meet the empty
        # strings the window writes past the year count and fail the whole cell.
        absolute_sum = (f'(SUMIF({profile_range},">0")-SUMIF({profile_range},"<0"))')
        cells = {
            # The authoritative TOTAL, taken from the cell this sheet already
            # publishes it in - not re-derived, and not the contingency below it.
            "total": f'=${display_col}${selected["quantile_row"]}',
            "profile_sum": (f'=IF({distribution_cell}="{at.not_produced}","",'
                            f'SUM({profile_range}))'),
            "difference": f'=IF({blank_guard},"",{total}-{profile_sum})',
            "allowance": (f'=IF({blank_guard},"",MAX({absolute},{relative}*'
                          f'MAX({scale_floor},ABS({total})+{absolute_sum})))'),
            "status": (f'=IF(OR({total}="",{difference}=""),"",'
                       f'IF(ABS({difference})<={allowance},'
                       f'"{verdicts["reconciled"]}","{verdicts["mismatch"]}")'
                       f'&IF({profile_state_cell}="{at.profile_current}",""'
                       f',"{verdicts["qualifier_prefix"]}"&{profile_state_cell}&'
                       f'"{verdicts["qualifier_suffix"]}"))'),
        }
        for entry in reconciliation["rows"]:
            key = entry["key"]
            row = rows[key]
            if measure == "nominal":
                _write(worksheet, f"{label_col}{row}", entry["label"], styles.label)
            cell = worksheet[f"{display_col}{row}"]
            cell.value = cells[key]
            cell.font = styles.value
            cell.number_format = formats[entry["format"]]


def _excel_number(value: float) -> str:
    """A contract tolerance, written so Excel parses it back to the same Double.

    PLAIN DECIMAL, NOT `repr`. `repr(1e-12)` is `1e-12`, and a lowercase
    exponent in a formula is a bet on a parser this project cannot test from
    Linux. `0.000000000001` reads back to the same binary64 and cannot be
    misparsed by anything.
    """
    text = format(Decimal(repr(float(value))), "f")
    if Decimal(text) != Decimal(repr(float(value))):  # pragma: no cover - defensive
        raise ValueError(f"{value!r} could not be written as an exact decimal")
    return text


def _populate_blocks(
    worksheet: Worksheet,
    sheet_spec: SheetSpec,
    styles: StyleBook,
    metadata: BuildMetadata,
) -> None:
    """Sheets whose body comes from the manifest rather than the input contract."""
    layout = styles.layout
    label_col = layout.label_column
    value_col = layout.value_column

    row = layout.body_start_row
    for block in sheet_spec.blocks:
        row = _write_block(worksheet, block, row, styles, metadata, label_col, value_col)
    return None


def _rule_columns(sheet_spec: SheetSpec) -> list[str]:
    if not sheet_spec.column_widths:
        return ["B"]
    return sorted(sheet_spec.column_widths, key=lambda c: (len(c), c))


def _write_block(
    worksheet: Worksheet,
    block: dict[str, Any],
    row: int,
    styles: StyleBook,
    metadata: BuildMetadata,
    label_col: str,
    value_col: str,
) -> int:
    if block["type"] == "note":
        _write(worksheet, f"{label_col}{row}", block["text"], styles.note)
        return row + 2

    _write(worksheet, f"{label_col}{row}", block["title"], styles.section)
    worksheet.row_dimensions[row].height = styles.row_height("section")
    row += 1

    if block.get("note"):
        _write(worksheet, f"{label_col}{row}", block["note"], styles.note)
        row += 1

    for entry in block.get("rows") or []:
        _write(worksheet, f"{label_col}{row}", entry["label"], styles.label)
        if "value" in entry and entry["value"] is not None:
            font = styles.value_locked if entry.get("locked") else styles.value
            _write(worksheet, f"{value_col}{row}", entry["value"], font)
        if entry.get("note"):
            _write(worksheet, f"F{row}", entry["note"], styles.note)
        row += 1

    for item in block.get("list") or []:
        _write(worksheet, f"{value_col}{row}", item, styles.list_item)
        row += 1

    if block.get("metadata"):
        for label, value in metadata.as_rows():
            _write(worksheet, f"{label_col}{row}", label, styles.label)
            _write(worksheet, f"{value_col}{row}", value, styles.value)
            row += 1

    return row + 1


def _write(worksheet: Worksheet, address: str, value: Any, font) -> None:
    cell = worksheet[address]
    cell.value = value
    cell.font = font
