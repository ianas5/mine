#!/usr/bin/env python3
"""The Phase-6 Gate-B inspection projection: `build/phase6_gate_b_inspection.json`.

WHY THIS FILE EXISTS
--------------------
The Windows Gate-B harness has to find things in the driven workbook before it
can say anything about a simulation: the `_SimData` sheet, the twenty-three run
identity rows and what each one MEANS, which column is bank A and which is bank
B, the durable pending-AUTO-nonce sidecar, the single final-commit range, the
iteration/summary/contingency geometry it actually inspects, and the two Setup
controls a fixture writes.

None of that is projected by any existing build output:

  * `stage_b_manifest.json` projects sheets, modules, entry points, API
    procedures, buttons, defined names for the timeline and counters, the two
    driver registers and the three grids - and stops there. No `_SimData`
    geometry, no `F21`, no simulation control;
  * `phase5_gate_b_inspection.json` is the PHASE-5 projection. It carries the
    `_Calc` layout and the Phase-5 fixture inputs, and materially broadening it
    to carry a second phase's machine sheet would make one artefact answer for
    two contracts;
  * `phase6_cases.json` is an expected-VALUE corpus. Its `J_publication` group
    carries the publication STATE MACHINE - `active_bank`, `next_auto_nonce`,
    `last_run_id`, `last_attempt_result` - and not one address. `F21` does not
    appear in it at all.

The addresses do exist in `build/vba/modSimContract.bas` and
`build/vba/modConstants.bas`, but those are VBA SOURCE. Teaching PowerShell to
parse VBA constants would put a second reader of the same layout authority in
the harness, which is exactly what "a second contract" means in practice.

WHAT THIS IS AND IS NOT
-----------------------
A PROJECTION, not a contract. Every value is read from the already-accepted
authorities - `sim_contract.yaml`, `input_contract.yaml` and `workbook.yaml` -
through their own loaders, and nothing is restated here.

IDENTITIES ONLY: sheets, columns, rows, ranges, cells, defined names, and the
public procedure names the contract already settles. It carries no expected
number, no tolerance, no bound and no vocabulary:

  * `SIM_MIN_ITERATIONS`, the seed domain and the run-ID maximum are VALUES that
    a scenario compares against, so they belong to the expectation corpus
    (`phase6_gate_b_cases.json`), not to the thing that says where to look;
  * the attempt-result and simulation-status label sets are model SEMANTICS. The
    Phase-5 projection had `attempt_result_labels` in its first submission and
    independent review removed them for exactly that reason. The same line is
    held here.
  * the three failpoint stage names are declared in PRODUCTION VBA
    (`modSimNonce.bas`, `modSimReport.bas`), not in a contract, so they are not
    projected. The harness declares them once and
    `tests/test_phase6_gate_b_harness_source.py` pins those strings against the
    modules that own them.

`tests/test_phase6_gate_b_harness_source.py` pins every address emitted here
against the generated `modSimContract.bas` and `modConstants.bas`, and the
sheet against the generated Stage-A workbook itself, so the two projections
cannot drift: if they ever disagree the Linux build fails, rather than the
harness silently inspecting the wrong cell on Windows.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .contract_loader import InputContract
from .sim_loader import SimContract

SCHEMA_VERSION = 1

INSPECTION_FILENAME = "phase6_gate_b_inspection.json"

# THE POSITIVE SCHEMA. Every key this projection may carry, at every level,
# named here. A banned-name list can only refuse the semantic values somebody
# already thought of; an allowlist refuses the next one too, whatever it is
# called. `tests/…_validation.py` plants an extra key and requires refusal.
ALLOWED_ROOT_KEYS = ("schema_version", "purpose", "provenance", "sim_data",
                     "publication", "controls", "command_surface")
ALLOWED_PROVENANCE_KEYS = ("sim_contract", "sim_contract_version",
                           "input_contract", "input_contract_version")
ALLOWED_SIM_DATA_KEYS = ("sheet", "required_visibility", "run_identity",
                         "pending_auto_nonce", "iteration_records",
                         "summary_statistics", "contingency_ladder")
ALLOWED_IDENTITY_KEYS = ("label_column", "value_column", "note_column",
                         "bank_value_columns", "first_row", "last_row", "rows",
                         "groups", "labels")
ALLOWED_PENDING_KEYS = ("cell", "column", "row", "label")
ALLOWED_ITERATION_KEYS = ("header_row", "first_iteration_row", "footer_rows",
                          "columns", "banks")
ALLOWED_BLOCK_KEYS = ("label_column", "bank_value_columns", "first_row",
                      "last_row", "rows")
ALLOWED_PUBLICATION_KEYS = ("bank_labels", "candidate_target",
                            "final_commit_range", "final_commit_fields")
ALLOWED_CONTROL_KEYS = ("defined_name", "sheet", "cell", "type")
ALLOWED_SURFACE_KEYS = ("automation_endpoint", "read_accessors")

# The two Setup scalars a Phase-6 scenario writes. Named by CONTRACT KEY; the
# defined name, sheet, cell and type all come from the contract itself.
CONTROL_INPUT_KEYS = ("monte_carlo_iterations", "random_seed")


@dataclass(frozen=True)
class SimInspectionArtifact:
    path: Path


def build_sim_inspection(sim: SimContract, contract: InputContract) -> dict[str, Any]:
    """The projection, as plain data. Identities only."""
    return {
        "schema_version": SCHEMA_VERSION,
        "purpose": (
            "Inspection identities for the Phase-6 Gate-B Windows harness. "
            "Addresses and names only, projected from the accepted contracts. "
            "Expected values live in phase6_gate_b_cases.json and phase6_cases.json."
        ),
        "provenance": {
            "sim_contract": sim.source_path.name,
            "sim_contract_version": sim.version,
            "input_contract": contract.source_path.name,
            "input_contract_version": contract.contract_version,
        },
        "sim_data": _sim_data_projection(sim),
        "publication": _publication_projection(sim),
        "controls": _control_projection(contract),
        "command_surface": _command_surface_projection(sim),
    }


def _sim_data_projection(sim: SimContract) -> dict[str, Any]:
    block = sim.raw["sim_data"]
    identity = block["run_identity"]
    fields = identity["fields"]
    iteration = block["iteration_records"]
    pending = block["pending_auto_nonce"]
    return {
        # The sheet and its visibility come from the layout the loader parsed,
        # not from the raw mapping, so a projection cannot disagree with the
        # value the rest of the build already used.
        "sheet": sim.layout.sheet,
        "required_visibility": sim.layout.required_visibility,
        "run_identity": {
            "label_column": identity["label_column"],
            "value_column": identity["value_column"],
            "note_column": identity["note_column"],
            "bank_value_columns": dict(identity["bank_value_columns"]),
            "first_row": identity["first_row"],
            "last_row": identity["last_row"],
            # ROW BY MEANING. The harness never counts rows and never adds an
            # offset: it asks for `next_auto_nonce` and is told 21.
            "rows": {field["key"]: field["row"] for field in fields},
            # WHICH AXIS EACH ROW BELONGS TO. `snapshot` rows are per-bank,
            # `counter` and `attempt` rows are shared. A scenario that captured
            # a shared row from a bank column would be reading the wrong cell,
            # and this is what stops it having to know that by heart.
            "groups": {field["key"]: field["group"] for field in fields},
            "labels": {field["key"]: field["label"] for field in fields},
        },
        "pending_auto_nonce": {
            "cell": pending["cell"],
            "column": pending["column"],
            "row": pending["row"],
            "label": pending["label"],
        },
        "iteration_records": {
            "header_row": sim.layout.header_row,
            "first_iteration_row": sim.layout.first_iteration_row,
            "footer_rows": sim.layout.footer_rows,
            "columns": [column["key"] for column in iteration["columns"]],
            "banks": {
                bank: dict(columns) for bank, columns in iteration["banks"].items()
            },
        },
        "summary_statistics": _measure_block(block["summary_statistics"], "metrics"),
        "contingency_ladder": _measure_block(block["contingency_ladder"], "rungs"),
    }


def _measure_block(block: dict[str, Any], row_key: str) -> dict[str, Any]:
    """A per-measure, per-bank value block: two columns per bank, rows by key."""
    return {
        "label_column": block["label_column"],
        "bank_value_columns": {
            bank: dict(columns)
            for bank, columns in block["bank_value_columns"].items()
        },
        "first_row": block["first_row"],
        "last_row": block["last_row"],
        "rows": {entry["key"]: entry["row"] for entry in block[row_key]},
    }


def _publication_projection(sim: SimContract) -> dict[str, Any]:
    publication = sim.raw["publication"]
    banks = publication["banks"]
    transaction = publication["transaction"]
    return {
        "bank_labels": list(banks["labels"]),
        # THE SELECTOR MAP, not a rule the harness restates. `""` is the blank
        # active-bank cell of a workbook that has never published.
        "candidate_target": dict(banks["candidate_target"]),
        "final_commit_range": transaction["final_commit_range"],
        "final_commit_fields": list(transaction["final_commit_fields"]),
    }


def _control_projection(contract: InputContract) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for key in CONTROL_INPUT_KEYS:
        spec = contract.inputs[key]
        out[key] = {
            "defined_name": spec.defined_name,
            "sheet": spec.sheet,
            "cell": spec.cell,
            "type": spec.type,
        }
    return out


def _command_surface_projection(sim: SimContract) -> dict[str, Any]:
    surface = sim.raw["command_surface"]
    return {
        "automation_endpoint": surface["automation_endpoint"],
        "read_accessors": list(surface["read_accessors"]),
    }


def emit_sim_inspection(
    build_dir: Path, sim: SimContract, contract: InputContract
) -> SimInspectionArtifact:
    """Write `phase6_gate_b_inspection.json`. Returns its path."""
    build_dir = Path(build_dir)
    build_dir.mkdir(parents=True, exist_ok=True)
    path = build_dir / INSPECTION_FILENAME
    path.write_text(
        json.dumps(build_sim_inspection(sim, contract), indent=2, sort_keys=False)
        + "\n",
        encoding="utf-8",
    )
    return SimInspectionArtifact(path=path)
