#!/usr/bin/env python3
"""The Gate-B inspection projection: `build/phase5_gate_b_inspection.json`.

WHY THIS FILE EXISTS
--------------------
The Windows Gate-B harness has to find things in the driven workbook: the five
`_Calc` ListObjects, the `calc_state` and `calc_totals` value cells and what each
row MEANS, the Setup scalars a fixture writes, and the Config/Setup tables a
fixture seeds. None of those identities is projected by any existing build
output:

  * `stage_b_manifest.json` projects sheets, modules, entry points, API
    procedures, buttons, the timeline/counter defined names, the two driver
    registers and the three grids - and stops there. The `_Calc` layout, the
    Setup input scalars and the Config lookup tables are absent. What the
    manifest DOES project is not repeated here: the harness reads registers and
    grids from the manifest, exactly as the Phase-4 harness already does;
  * `phase5_cases.json` is an expected-VALUE corpus. It carries no addresses;
  * `phase4_scenarios.json` is the structural oracle's output, also value-only.

The layout does exist in `build/vba/modCalcContract.bas` and
`build/vba/modConstants.bas`, but those are VBA source. Teaching PowerShell to
parse VBA constants would put a second reader of the same authority in the
harness, which is exactly what a "second contract" means in practice.

WHAT THIS IS AND IS NOT
-----------------------
This is a PROJECTION, not a contract. Every value here is read from the three
already-accepted authorities - `calc_contract.yaml`, `input_contract.yaml` and
`driver_contract.yaml` - through their own loaders, and nothing is restated. It
carries IDENTITIES ONLY: names, sheets, addresses, rows, columns. It carries no
expected numerical value, no tolerance and no analytical fact; those stay in
`phase5_cases.json`, which remains the sole expected-value authority.

`tests/test_phase5_gate_b_harness_source.py` pins every address here against the
generated `modCalcContract.bas` and `modConstants.bas`, so the two projections
cannot drift apart: if they ever disagree, the build fails on Linux rather than
the harness silently inspecting the wrong cell on Windows.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .calc_loader import CalcContract
from .contract_loader import InputContract

SCHEMA_VERSION = 1

INSPECTION_FILENAME = "phase5_gate_b_inspection.json"

# The Setup scalars a Gate-B fixture writes or reads. Named by contract KEY; the
# defined name, sheet and cell all come from the contract itself.
FIXTURE_INPUT_KEYS = (
    "duration_years",
    "project_start_year",
    "base_year",
    "discount_rate",
    "selected_confidence_level",
    "reporting_currency",
    "project_name",
)

# The lookup/rate tables a Gate-B fixture seeds. Same rule: keys only.
FIXTURE_TABLE_KEYS = (
    "fx_rates",
    "currencies",
    "categories",
    "uom",
    "inflation_profiles",
    "distributions",
    "confidence_levels",
)


@dataclass(frozen=True)
class InspectionArtifact:
    path: Path


def build_inspection(calc: CalcContract, contract: InputContract) -> dict[str, Any]:
    """The projection, as plain data. Identities only."""
    return {
        "schema_version": SCHEMA_VERSION,
        "purpose": (
            "Inspection identities for the Phase-5 Gate-B Windows harness. "
            "Addresses and names only, projected from the accepted contracts. "
            "Expected values live in phase5_cases.json and nowhere else."
        ),
        "provenance": {
            "calc_contract": str(calc.source_path.name),
            "calc_contract_version": calc.version,
            "input_contract": str(contract.source_path.name),
            "input_contract_version": contract.contract_version,
        },
        "calc": _calc_projection(calc),
        "inputs": _input_projection(contract),
        "input_tables": _input_table_projection(contract),
    }


def _calc_projection(calc: CalcContract) -> dict[str, Any]:
    return {
        "sheet": calc.sheet,
        "required_visibility": calc.required_visibility,
        "fingerprint_version": calc.fingerprint_version,
        "derived_status_labels": list(calc.derived_status_labels),
        "attempt_result_labels": list(calc.attempt_result_labels),
        "tables": {
            key: {
                "table_name": table.table_name,
                "header_row": table.header_row,
                "first_column": table.first_column,
                "last_column": table.last_column,
                "first_column_index": table.first_column_index,
                "column_count": table.band_width,
                "first_body_row": table.header_row + 1,
                "row_rule": table.row_rule,
                "columns": [column.key for column in table.columns],
            }
            for key, table in calc.tables.items()
        },
        "scalar_blocks": {
            key: {
                "label_column": block.label_column,
                "value_column": block.value_column,
                "first_row": block.first_row,
                "last_row": block.last_row,
                "value_range": block.value_range(),
                "rows": {field.key: field.row for field in block.fields},
            }
            for key, block in calc.scalar_blocks.items()
        },
    }


def _input_projection(contract: InputContract) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for key in FIXTURE_INPUT_KEYS:
        spec = contract.inputs[key]
        out[key] = {
            "defined_name": spec.defined_name,
            "sheet": spec.sheet,
            "cell": spec.cell,
            "type": spec.type,
        }
    return out


def _input_table_projection(contract: InputContract) -> dict[str, Any]:
    by_key = {table.key: table for table in contract.all_tables}
    out: dict[str, Any] = {}
    for key in FIXTURE_TABLE_KEYS:
        table = by_key[key]
        out[key] = {
            "table_name": table.table_name,
            "sheet": table.sheet,
            "header_row": table.header_row,
            "first_column": table.first_column,
            "columns": [column.header for column in table.columns],
            "locked_seed_rows": table.locked_seed_rows,
        }
    return out


def emit_inspection(
    build_dir: Path, calc: CalcContract, contract: InputContract
) -> InspectionArtifact:
    """Write `phase5_gate_b_inspection.json`. Returns its path."""
    build_dir = Path(build_dir)
    build_dir.mkdir(parents=True, exist_ok=True)
    path = build_dir / INSPECTION_FILENAME
    path.write_text(
        json.dumps(build_inspection(calc, contract), indent=2, sort_keys=False) + "\n",
        encoding="utf-8",
    )
    return InspectionArtifact(path=path)
