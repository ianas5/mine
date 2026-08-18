"""Emit the Stage-B inputs from the contracts.

Two generated artifacts, both written into the build directory beside the Stage-A
workbook:

  build/vba/modConstants.bas    every structural literal the VBA needs
  build/stage_b_manifest.json   every structural literal the PowerShell needs
  build/phase4_scenarios.json   the timeline shapes the functional harness asserts

Neither is hand-maintained. The structure contract says a value once; this module
projects it into the two languages that cannot read YAML. That is what keeps the
contract's "do not scatter these values through VBA, PowerShell or Python" promise
honest: the bootstrap and the harness read the manifest instead of restating sheet
names, CodeNames, button captions, macro names or module lists.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from .contract_loader import InputContract
from .driver_loader import DriverContract
from .scenarios import build_scenarios
from .spec_loader import WorkbookSpec
from .structure_loader import StructureContract
from .structure_oracle import Limits as OracleLimits

GENERATED_BANNER = (
    "GENERATED FILE - DO NOT EDIT.\n"
    "Emitted from spec/structure_contract.yaml by the Stage-A builder.\n"
    "Edit the contract and rebuild; edits made here are overwritten."
)

XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52

VBA_LONG_MAX = 2_147_483_647
"""The largest value a VBA Long can hold.

Emitted so the VBA can bound an untrusted numeric cell BEFORE converting it. It is
a representational limit of the language, never a limit on the model: identifiers,
years and durations all have their own contract-declared bounds.
"""


@dataclass(frozen=True)
class StageBArtifacts:
    module_path: Path
    manifest_path: Path
    scenario_path: Path


# ---------------------------------------------------------------------------
def emit_stage_b(
    build_dir: Path,
    spec: WorkbookSpec,
    contract: InputContract,
    drivers: DriverContract,
    structure: StructureContract,
) -> StageBArtifacts:
    """Write modConstants.bas and stage_b_manifest.json. Returns their paths."""
    build_dir = Path(build_dir)
    vba_dir = build_dir / Path(structure.vba_generated_dir).name
    vba_dir.mkdir(parents=True, exist_ok=True)

    module_path = vba_dir / f"{structure.vba_generated_module}.bas"
    module_path.write_text(
        render_constants_module(spec, contract, drivers, structure), encoding="utf-8"
    )

    manifest_path = build_dir / "stage_b_manifest.json"
    manifest_path.write_text(
        json.dumps(
            build_manifest(spec, contract, drivers, structure), indent=2, sort_keys=False
        )
        + "\n",
        encoding="utf-8",
    )

    scenario_path = build_dir / "phase4_scenarios.json"
    scenario_path.write_text(
        json.dumps(build_scenarios(oracle_limits(structure)), indent=2, sort_keys=False) + "\n",
        encoding="utf-8",
    )
    return StageBArtifacts(
        module_path=module_path,
        manifest_path=manifest_path,
        scenario_path=scenario_path,
    )


def oracle_limits(structure: StructureContract) -> OracleLimits:
    """The contract's limits, as the oracle's own value type.

    One conversion point, so the oracle never reads YAML and the contract never
    imports the oracle.
    """
    return OracleLimits(
        min_year=structure.limits.min_year,
        max_year=structure.limits.max_year,
        max_generated_year_columns=structure.limits.max_generated_year_columns,
    )


# ---------------------------------------------------------------------------
# modConstants.bas
# ---------------------------------------------------------------------------
def render_constants_module(
    spec: WorkbookSpec,
    contract: InputContract,
    drivers: DriverContract,
    structure: StructureContract,
) -> str:
    lines: list[str] = ["Attribute VB_Name = \"" + structure.vba_generated_module + "\""]
    lines.append("Option Explicit")
    lines.append("")
    for banner_line in GENERATED_BANNER.splitlines():
        lines.append(f"' {banner_line}")
    lines.append("'")
    lines.append(f"' Model version    : {spec.model['model_version']}")
    lines.append(f"' Structure contract: {structure.version}")
    lines.append("")

    def section(title: str) -> None:
        lines.append("' " + "-" * 72)
        lines.append(f"' {title}")
        lines.append("' " + "-" * 72)

    def const(name: str, value, comment: str | None = None) -> None:
        rendered = _vba_literal(value)
        if isinstance(value, str):
            kind = "String"
        elif isinstance(value, float):
            # A Long literal of `1.0` does not parse. A rate is a Double.
            kind, rendered = "Double", _vba_double(value)
        else:
            kind = "Long"
        text = f"Public Const {name} As {kind} = {rendered}"
        if comment:
            text += f"    ' {comment}"
        lines.append(text)

    # --- sheets ------------------------------------------------------------
    section("Sheet names")
    for sheet in spec.sheets:
        const(f"SH_{_ident(sheet.name)}", sheet.name)
    lines.append("")

    # --- tables ------------------------------------------------------------
    section("Table names")
    for register in drivers.all_registers:
        const(f"TBL_{_ident(register.key)}", register.table_name)
    for grid in structure.all_grids:
        const(f"TBL_{_ident(grid.key)}", grid.table_name)
    for table in contract.all_tables:
        const(f"TBL_{_ident(table.key or table.table_name)}", table.table_name)
    lines.append("")

    # --- Setup / Config table geometry the CALCULATION reads ----------------
    # Phase 4 manages structure and never reads a rate, a discount or a
    # distribution name, so these were not projected before. Phase 5 resolution
    # does read them, and the alternative to projecting them is a second copy of
    # the contract's own coordinates hand-written into VBA.
    section("Setup / Config table geometry")
    for table in contract.all_tables:
        prefix = _ident(table.key or table.table_name)
        for ordinal, column in enumerate(table.columns, start=1):
            const(f"COL_{prefix}_{_ident(column.header)}", ordinal, column.header)
        const(f"TBL_{prefix}_LOCKED_SEED_ROWS", table.locked_seed_rows)
    lines.append("")

    # --- input defined names ------------------------------------------------
    # The Setup input cells, by defined name. Phase 5 reads the discount rate;
    # the rest are projected with it so the set has one rule rather than an
    # ad-hoc membership decided by whichever field a later phase happened to
    # need first.
    section("Setup input defined names")
    for key in sorted(contract.inputs):
        const(f"NM_INPUT_{_ident(key)}", contract.inputs[key].defined_name)
    lines.append("")

    # --- locked model vocabulary -------------------------------------------
    # The reporting currency and its identity rate come from the FX table's own
    # locked seed row, and the distribution master list from the Config table
    # that owns it. Neither is restated here or in VBA: the NAMES are projected,
    # and which internal shape each distribution selects is an adapter that
    # belongs to the resolver, exactly as it belongs to the Python oracle.
    section("Locked model vocabulary")
    _project_reporting_currency(contract, const)
    distributions = _seeded_values(contract, "distributions")
    const("DISTRIBUTION_COUNT", len(distributions))
    for ordinal, name in enumerate(distributions, start=1):
        const(f"DISTRIBUTION_NAME_{ordinal}", name)
    lines.append("")

    # --- entered / applied names -------------------------------------------
    section("Timeline defined names (entered aliases, applied state, derived state)")
    for alias in structure.entered_aliases:
        const(f"NM_{_ident(alias.input_key)}_ENTERED", alias.defined_name)
    for field in structure.applied:
        const(f"NM_{_ident(field.key)}", field.defined_name)
    for field in structure.derived:
        const(f"NM_{_ident(field.key)}", field.defined_name)
    const("NM_STRUCTURAL_STATE", structure.structural_state.defined_name)
    lines.append("")

    section("Structural state vocabulary")
    for key, label in structure.structural_state.labels.items():
        const(f"STATE_{key.upper()}", label)
    lines.append("")

    # --- identity ----------------------------------------------------------
    section("Permanent identity")
    # An alias, not a second copy: the counters live on a sheet already named above.
    lines.append(f"Public Const SH_IDENTITY As String = SH_{_ident(structure.identity_sheet)}")
    for counter in structure.counters:
        upper = _ident(counter.key)
        const(f"NM_COUNTER_{upper}", counter.defined_name)
        const(f"ID_PREFIX_{upper}", counter.prefix)
        const(f"ID_PAD_{upper}", counter.pad_width, "minimum display width, not a maximum")
    lines.append("' The largest sequence a VBA Long can represent.")
    lines.append("'")
    lines.append("' This is an IMPLEMENTATION REPRESENTATION CEILING, not a business maximum.")
    lines.append("' The model imposes no limit on how many identifiers may be issued, but this")
    lines.append("' implementation cannot represent a sequence beyond this value, so allocation")
    lines.append("' refuses CLEANLY at the ceiling rather than overflowing, and a stored counter")
    lines.append("' or an ID tail beyond it is reported as corrupt rather than silently ignored.")
    lines.append("'")
    lines.append("' A counter sitting exactly AT this value is VALID, exhausted state: no further")
    lines.append("' identifier can be allocated, but the existing structure stays sound and")
    lines.append("' structural revalidation must remain clean.")
    const("ID_COUNTER_MAX", VBA_LONG_MAX)
    lines.append("")

    # --- limits ------------------------------------------------------------
    section("Structural limits (generation guards, never business maxima)")
    lines.append("' Two INDEPENDENT protections. They guard different things:")
    lines.append("'   LIMIT_MIN_YEAR / LIMIT_MAX_YEAR   the supported CALENDAR-YEAR window,")
    lines.append("'     bounding Base Year, Start Year, Last Project Year and the inflation span.")
    lines.append("'   LIMIT_MAX_YEAR_COLUMNS            Architecture Lock Revision B protection on")
    lines.append("'     generated PROJECT-YEAR columns: \"Generated column count > 200 = ERROR\".")
    lines.append("' Neither is derived from the other, and neither is a business maximum.")
    const("LIMIT_MIN_YEAR", structure.limits.min_year)
    const("LIMIT_MAX_YEAR", structure.limits.max_year)
    const(
        "LIMIT_MAX_YEAR_COLUMNS",
        structure.limits.max_generated_year_columns,
        "generated project-year columns; NOT a calendar-year cap",
    )
    lines.append("")

    # --- driver registers ---------------------------------------------------
    section("Driver register geometry")
    for register in drivers.all_registers:
        upper = _ident(register.key)
        const(f"REG_{upper}_SHEET", register.sheet)
        const(f"REG_{upper}_HEADER_ROW", register.header_row)
        const(f"REG_{upper}_FIRST_COL", register.first_col_index)
        const(f"REG_{upper}_COL_COUNT", len(register.columns))
        const(f"REG_{upper}_ID_COL", register.column_index_of(register.columns[0].key) + 1)
        for index, column in enumerate(register.columns):
            const(f"COL_{upper}_{_ident(column.key)}", index + 1)
    lines.append("")

    # --- grids --------------------------------------------------------------
    section("Structural grid geometry")
    for grid in structure.all_grids:
        upper = _ident(grid.key)
        const(f"GRID_{upper}_SHEET", grid.sheet)
        const(f"GRID_{upper}_HEADER_ROW", grid.header_row)
        const(f"GRID_{upper}_FIRST_COL", grid.first_col_index)
        const(f"GRID_{upper}_FIXED_COLS", len(grid.fixed_columns))
        const(f"GRID_{upper}_RESERVED_ROWS", grid.reserved_rows)
        const(f"GRID_{upper}_YEAR_FORMAT", grid.year_column.number_format)
        const(f"GRID_{upper}_HEADER_FORMAT", grid.year_column.header_format)
        const(f"GRID_{upper}_YEAR_WIDTH", int(grid.year_column.width))
        for index, column in enumerate(grid.fixed_columns):
            const(f"GCOL_{upper}_{_ident(column.key)}", index + 1)
    lines.append("")
    lines.append(
        "' A new profiling cell is initialised to 0%. A new inflation cell is left"
    )
    lines.append("' BLANK: an escalation assumption the user never made must not be invented.")
    lines.append(
        f"Public Const PROFILE_INITIAL_VALUE As Double = "
        f"{structure.profiling_grids[0].year_column.initial_value}"
    )
    lines.append("")

    # --- presentation -------------------------------------------------------
    lines.append("' A year column generated at RUNTIME must carry the SAME editable-input")
    lines.append("' treatment as every other user-owned cell. Relying on Excel table-format")
    lines.append("' propagation next to a model-controlled fixed column is not deterministic")
    lines.append("' enough, so the fill is applied explicitly -- from the contract, never from")
    lines.append("' a colour written into the VBA.")
    const("ERROR_CELL_MARKER", "#ERROR!",
          "deterministic marker for a cell holding an Excel error value")
    lines.append("")

    section("Input-language fills, as Excel Interior.Color (BGR) values")
    const("FILL_INPUT", _bgr(spec.presentation["colors"]["input_fill"]),
          f'#{spec.presentation["colors"]["input_fill"]}')
    const("FILL_LOCKED", _bgr(spec.presentation["colors"]["locked_fill"]),
          f'#{spec.presentation["colors"]["locked_fill"]}')
    lines.append("")

    # --- messages -----------------------------------------------------------
    section("Structural messages")
    for key in ("not_applied", "inflation_empty_span"):
        const(f"MSG_{key.upper()}", structure.state_messages[key])
    lines.append("")

    # --- entry points -------------------------------------------------------
    section("Public entry points bound to the command buttons")
    for name in structure.entry_points:
        const(f"ENTRY_{_ident(name.replace('PCCM_', ''))}", name)
    lines.append("")

    section("Structural check keys reported by modStructuralCheck")
    for check in structure.structural_checks:
        const(f"CHK_{_ident(check['key'])}", check["key"])
    lines.append("")

    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# stage_b_manifest.json
# ---------------------------------------------------------------------------
def build_manifest(
    spec: WorkbookSpec,
    contract: InputContract,
    drivers: DriverContract,
    structure: StructureContract,
) -> dict:
    """Everything the Windows bootstrap and functional harness need, as plain data."""
    return {
        "generated_by": "pccm stage-a builder",
        "model_version": spec.model["model_version"],
        "manifest_version": spec.manifest_version,
        "structure_contract_version": structure.version,
        "stage_a_filename": spec.stage_a_filename,
        "stage_b_filename": spec.stage_a_filename.replace("_stageA.xlsx", "_stageB.xlsm"),
        "xlsm_file_format": XL_OPEN_XML_WORKBOOK_MACRO_ENABLED,
        "sheets": [
            {"name": s.name, "codename": s.codename, "visibility": s.visibility}
            for s in spec.sheets
        ],
        "vba": {
            "source_dir": structure.vba_source_dir,
            "generated_dir": structure.vba_generated_dir,
            "modules": [
                {"name": m.name, "generated": m.generated, "responsibility": m.responsibility}
                for m in structure.vba_modules
            ],
            "entry_points": list(structure.entry_points),
            # The calculation endpoints a later harness drives through
            # Application.Run. Deliberately NOT entry_points: nothing binds them
            # to a button, and the manifest is where the harness learns the
            # difference.
            "api_procedures": list(structure.api_procedures),
            "forbidden_constructs": list(structure.forbidden_constructs),
        },
        "buttons": [
            {
                "key": b.key,
                "sheet": b.sheet,
                "shape_name": b.shape_name,
                "caption": b.caption,
                "entry_point": b.entry_point,
                "anchor_cell": b.anchor_cell,
                "width": b.width,
                "height": b.height,
            }
            for b in structure.buttons
        ],
        "defined_names": {
            **structure.defined_names,
            **structure.alias_defined_names(contract),
        },
        "counters": [
            {
                "key": c.key,
                "defined_name": c.defined_name,
                "cell": c.cell,
                "sheet": structure.identity_sheet,
                "prefix": c.prefix,
                "pad_width": c.pad_width,
                "pattern": c.pattern,
                "driver_register": c.driver_register,
                "initial": c.initial,
            }
            for c in structure.counters
        ],
        "registers": [
            {
                "key": r.key,
                "sheet": r.sheet,
                "table_name": r.table_name,
                "header_row": r.header_row,
                "first_column": r.first_column,
                "columns": [c.key for c in r.columns],
                "reserved_rows": r.reserved_rows,
            }
            for r in drivers.all_registers
        ],
        "grids": [
            {
                "key": g.key,
                "kind": g.kind,
                "sheet": g.sheet,
                "table_name": g.table_name,
                "header_row": g.header_row,
                "first_column": g.first_column,
                "fixed_columns": [c.key for c in g.fixed_columns],
                "reserved_rows": g.reserved_rows,
                "driver_register": g.driver_register,
                "year_number_format": g.year_column.number_format,
                "year_initial_value": g.year_column.initial_value,
            }
            for g in structure.all_grids
        ],
        "limits": {
            "min_year": structure.limits.min_year,
            "max_year": structure.limits.max_year,
            "max_generated_year_columns": structure.limits.max_generated_year_columns,
            # The representation ceiling, from the SAME source as the VBA constant, so
            # the functional harness can drive a counter to it without restating the
            # number. It is emitted as a limit, not as a business maximum: a counter
            # sitting at it is valid, exhausted state.
            "id_counter_max": VBA_LONG_MAX,
        },
        "state_labels": dict(structure.structural_state.labels),
        # Presentation tokens, so the functional harness can assert that a row or a
        # year column CREATED AT RUNTIME still carries the Phase-2 input language --
        # editable cells must never become visually indistinguishable from
        # model-controlled ones. The harness restates no colour of its own.
        "presentation": {
            "input_fill": spec.presentation["colors"]["input_fill"],
            "locked_fill": spec.presentation["colors"]["locked_fill"],
        },
        "driver_validation_columns": {
            register.key: [
                column.key
                for column in register.columns
                if column.validation is not None
            ]
            for register in drivers.all_registers
        },
        "structural_checks": [dict(c) for c in structure.structural_checks],
    }


# ---------------------------------------------------------------------------
def _vba_double(value: float) -> str:
    """A Double literal VBA parses back to the same bits.

    `repr` is the shortest decimal that round-trips and VBA's parser is correctly
    rounded, so the bits survive. An integral value gets an explicit `.0` so VBA
    cannot type it as a Long.
    """
    text = repr(float(value))
    if "e" in text or "E" in text or "." in text:
        return text
    return text + ".0"


def _seeded_values(contract: InputContract, key: str) -> list[str]:
    """The single-column seed values of a Config table, in contract order."""
    table = next(t for t in contract.all_tables if t.key == key)
    return [str(row[0]) for row in table.seed_rows]


def _project_reporting_currency(contract: InputContract, const) -> None:
    """The reporting currency and its identity rate, from the FX table itself.

    `tblFXRates` carries exactly one locked seed row, and that row IS the
    reporting-currency identity. Reading it here rather than naming the currency
    in this file keeps the currency and its rate where the contract already put
    them, and makes a change to that row a change to the projection.
    """
    table = next(t for t in contract.all_tables if t.table_name == "tblFXRates")
    if table.locked_seed_rows != 1 or len(table.seed_rows) != 1:
        raise ValueError(
            "the FX table must carry exactly one locked seed row: it is the "
            f"reporting-currency identity, found {len(table.seed_rows)}"
        )
    currency, rate = table.seed_rows[0][0], table.seed_rows[0][1]
    const("REPORTING_CURRENCY", str(currency))
    const("REPORTING_CURRENCY_RATE", float(rate),
          "a global invariant, enforced in every model")


def _bgr(hex_rgb: str) -> int:
    """Excel Interior.Color is a BGR long, not an RGB hex string."""
    value = hex_rgb.lstrip("#")
    red, green, blue = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return (blue << 16) | (green << 8) | red


def _ident(text: str) -> str:
    """A VBA-safe upper-snake identifier fragment."""
    out = []
    previous_underscore = False
    for char in text:
        if char.isalnum():
            out.append(char.upper())
            previous_underscore = False
        elif not previous_underscore:
            out.append("_")
            previous_underscore = True
    return "".join(out).strip("_")


def _vba_literal(value) -> str:
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, str):
        return '"' + value.replace('"', '""') + '"'
    return str(value)
