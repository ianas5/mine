"""Emit the Phase-5 generated artifacts.

Two files, written beside the Stage-A workbook:

  build/vba/modCalcContract.bas   the runtime literals later VBA needs
  build/phase5_cases.json         the expected-value corpus a later Windows
                                  harness will assert against

NEITHER IS EXECUTED HERE, and neither is embedded in the Stage-A `.xlsx`. Step 3
produces them; the VBA implementation steps consume them.

--------------------------------------------------------------------------------
WHY A GENERATED CONSTANTS MODULE EXISTS AT ALL
--------------------------------------------------------------------------------
Later VBA has to know that `tblCalcDrivers` starts at column X, that
`calc_state` occupies rows 13 to 20, that the profiling tolerance is `1e-9` and
that the hash base is 131. Every one of those already has exactly one owner. The
alternative to projecting them is hardcoding them in VBA, which creates a second
authority that drifts silently and is only discovered on Windows.

So this module says nothing of its own. Each constant is READ from the authority
that owns it:

    geometry, labels, tolerances, FP_VERSION  ->  spec/calc_contract.yaml
    FP_BASE, moduli, initial states, tags     ->  builder/pccm_builder/calc_fingerprint.py

The fingerprint mathematics is deliberately NOT in the YAML. `calc_fingerprint.py`
is the reference implementation and owns the hash outright, so the generator
imports it rather than restating its numbers; a test proves the emitted values are
the module's own and would change if the module changed.

What is NOT here: the fingerprint recurrence, any analytical formula, any
calculation. This is a constants module.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from . import calc_fingerprint as fp
from .calc_cases import build_cases
from .calc_loader import CalcContract
from .calc_render import placeholder_row, table_ref
from .spec_loader import WorkbookSpec

CALC_MODULE_NAME = "modCalcContract"

GENERATED_BANNER = (
    "GENERATED FILE - DO NOT EDIT.\n"
    "Emitted from spec/calc_contract.yaml and builder/pccm_builder/calc_fingerprint.py\n"
    "by the Stage-A builder. Edit the authority and rebuild; edits made here are\n"
    "overwritten. This module declares CONSTANTS ONLY - no calculation, no hash\n"
    "recurrence, no analytical formula."
)


@dataclass(frozen=True)
class CalcArtifacts:
    module_path: Path
    cases_path: Path


def emit_calc_artifacts(
    build_dir: Path, spec: WorkbookSpec, calc: CalcContract
) -> CalcArtifacts:
    """Write modCalcContract.bas and phase5_cases.json. Returns their paths."""
    build_dir = Path(build_dir)
    vba_dir = build_dir / "vba"
    vba_dir.mkdir(parents=True, exist_ok=True)

    module_path = vba_dir / f"{CALC_MODULE_NAME}.bas"
    module_path.write_text(render_calc_contract_module(spec, calc), encoding="utf-8")

    cases_path = build_dir / "phase5_cases.json"
    cases_path.write_text(render_cases_json(spec, calc), encoding="utf-8")
    return CalcArtifacts(module_path=module_path, cases_path=cases_path)


# ---------------------------------------------------------------------------
# modCalcContract.bas
# ---------------------------------------------------------------------------
def render_calc_contract_module(spec: WorkbookSpec, calc: CalcContract) -> str:
    """The generated VBA constants module, as text.

    Deterministic: the only inputs are the two authorities and the manifest's
    model version. No timestamp, no path, no environment.
    """
    lines: list[str] = [f'Attribute VB_Name = "{CALC_MODULE_NAME}"', "Option Explicit", ""]
    for banner_line in GENERATED_BANNER.splitlines():
        lines.append(f"' {banner_line}")
    lines.append("'")
    lines.append(f"' Model version       : {spec.model['model_version']}")
    lines.append(f"' Calculation contract: {calc.version}")
    lines.append("")

    def section(title: str) -> None:
        lines.append("' " + "-" * 72)
        lines.append(f"' {title}")
        lines.append("' " + "-" * 72)

    def const(name: str, value: Any, comment: str | None = None) -> None:
        if isinstance(value, bool):
            raise TypeError("booleans are not projected as constants")
        if isinstance(value, str):
            kind, rendered = "String", '"' + value.replace('"', '""') + '"'
        elif isinstance(value, int):
            kind, rendered = "Long", str(value)
        elif isinstance(value, float):
            kind, rendered = "Double", _vba_double(value)
        else:
            raise TypeError(f"cannot project {value!r}")
        text = f"Public Const {name} As {kind} = {rendered}"
        if comment:
            text += f"    ' {comment}"
        lines.append(text)

    # --- the workspace itself -----------------------------------------------
    section("Calculation workspace")
    const("CALC_SHEET", calc.sheet)
    const("CALC_SHEET_VISIBILITY", calc.required_visibility)
    const("CALC_PHASE4_FIRST_ROW", calc.phase4_first_row,
          "Phase-4 territory: never written by Phase 5")
    const("CALC_PHASE4_LAST_ROW", calc.phase4_last_row)
    lines.append("")

    # --- fingerprint --------------------------------------------------------
    section("Calculation Input Fingerprint (mathematics from calc_fingerprint.py)")
    const("FP_VERSION", calc.fingerprint_version,
          "from the calculation contract; the stamp written on a successful commit")
    const("FP_BASE", fp.FP_BASE)
    const("FP_MOD_1", fp.FP_MOD_1)
    const("FP_MOD_2", fp.FP_MOD_2)
    const("FP_INIT_1", fp.FP_INIT_1)
    const("FP_INIT_2", fp.FP_INIT_2)
    const("FP_STREAM_TAG", fp.STREAM_TAG)
    for index, name in enumerate(fp.SECTION_ORDER, start=1):
        const(f"FP_SECTION_{index}", name)
    const("FP_TAG_TEXT", fp.TAG_TEXT)
    const("FP_TAG_NUMBER", fp.TAG_NUMBER)
    const("FP_TAG_INTEGER", fp.TAG_INTEGER)
    lines.append("")

    # --- scalar blocks ------------------------------------------------------
    for block in (calc.calc_state, calc.calc_totals):
        prefix = block.key.upper()
        section(f"{block.key} block geometry")
        const(f"{prefix}_LABEL_COLUMN", block.label_column)
        const(f"{prefix}_VALUE_COLUMN", block.value_column)
        const(f"{prefix}_NOTE_COLUMN", block.note_column)
        const(f"{prefix}_FIRST_ROW", block.first_row)
        const(f"{prefix}_LAST_ROW", block.last_row)
        const(f"{prefix}_VALUE_RANGE", block.value_range())
        for entry in block.fields:
            const(f"{prefix}_ROW_{_ident(entry.key)}", entry.row)
        lines.append("")

    # --- vocabulary ---------------------------------------------------------
    section("Two orthogonal state axes")
    for label in calc.derived_status_labels:
        const(f"CALC_STATUS_{_ident(label)}", label)
    for label in calc.attempt_result_labels:
        const(f"CALC_ATTEMPT_{_ident(label)}", label)
    lines.append("")

    # --- tables -------------------------------------------------------------
    for table in calc.all_tables:
        prefix = _ident(table.key)
        section(f"{table.table_name} geometry")
        const(f"TBL_{prefix}", table.table_name)
        const(f"TBL_{prefix}_HEADER_ROW", table.header_row)
        const(f"TBL_{prefix}_FIRST_COLUMN", table.first_column)
        const(f"TBL_{prefix}_LAST_COLUMN", table.last_column)
        const(f"TBL_{prefix}_FIRST_COLUMN_INDEX", table.first_column_index)
        const(f"TBL_{prefix}_COLUMN_COUNT", len(table.columns))
        const(f"TBL_{prefix}_EMPTY_REF", table_ref(table),
              "header row plus one physically blank body row")
        const(f"TBL_{prefix}_FIRST_BODY_ROW", placeholder_row(table))
        for ordinal, column in enumerate(table.columns, start=1):
            const(f"COL_{prefix}_{_ident(column.key)}", ordinal,
                  f"{column.header}")
        lines.append("")

    # --- tolerances ---------------------------------------------------------
    section("Tolerances")
    const("TOL_PROFILING_SUM_ABSOLUTE", calc.tolerances.profiling_sum_absolute)
    const("TOL_IDENTITY_ABSOLUTE_FLOOR", calc.tolerances.identity_absolute_floor)
    const("TOL_IDENTITY_RELATIVE_COEFFICIENT",
          calc.tolerances.identity_relative_coefficient)
    const("TOL_CONDITIONING_SCALE_FLOOR", calc.tolerances.conditioning_scale_floor)
    lines.append("")

    return "\n".join(lines) + "\n"


def _vba_double(value: float) -> str:
    """A Double literal VBA parses back to the same bits.

    `repr` gives the shortest decimal that round-trips, and VBA's parser is
    correctly rounded, so the bits survive. An integral value gets an explicit
    `.0` so VBA cannot type it as a Long.
    """
    text = repr(float(value))
    if "e" in text or "E" in text or "." in text:
        return text
    return text + ".0"


def _ident(text: str) -> str:
    out = []
    for char in text:
        out.append(char.upper() if (char.isalnum() or char == "_") else "_")
    return "".join(out)


# ---------------------------------------------------------------------------
# phase5_cases.json
# ---------------------------------------------------------------------------
def render_cases_json(spec: WorkbookSpec, calc: CalcContract) -> str:
    """The acceptance corpus, as deterministic JSON text.

    `sort_keys=False` keeps the authored order of each record, which is stable
    because the source structures are tuples and dicts built in a fixed order.
    `allow_nan=False` makes a non-finite value a build failure rather than the
    non-standard `NaN`/`Infinity` tokens no other JSON reader accepts.
    """
    document = build_cases(calc, spec.model["model_version"])
    return json.dumps(document, indent=2, sort_keys=False, allow_nan=False,
                      ensure_ascii=True) + "\n"
