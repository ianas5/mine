#!/usr/bin/env python3
"""The Phase-8 Results projection: WHERE the P8-1 output surface lives.

WHY A THIRD PROJECTION AND NOT A WIDER SECOND. `phase6_gate_b_inspection.json`
is the artefact Run 6 was accepted against and `phase7_acceptance_inspection.json`
is Phase 7's; both are historical evidence, and widening either to carry a
Phase-8 layout would change a file whose identity is the evidence. So Phase 8
gets its own, on exactly the terms Phase 7 got its own.

IDENTITIES ONLY. Sheet, columns, rows, procedure names, header text and number
formats - the same line the Phase-5, Phase-6 and Phase-7 projections hold. No
expected value, no tolerance of its own; the identity allowance below is
`calc_contract.yaml`'s, carried so a Windows runner can check the sheet against
the same rule the sheet was built from rather than against a number typed into
a harness.

WHAT IT IS FOR. A Windows runner must read the Results surface without spelling
a single address, for the reason P7-4 made expensive: a hand-written address is
a second declaration of something a manifest already owns, and it goes stale
silently. Every row, column and procedure a P8-1 scenario touches is here.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .artifact_io import write_lf_artifact
from .calc_loader import CalcContract
from .spec_loader import WorkbookSpec

SCHEMA_VERSION = 1

INSPECTION_FILENAME = "phase8_results_inspection.json"

ALLOWED_KEYS = ("schema_version", "purpose", "provenance", "sheet", "columns",
                "run_stamp", "summary", "selected", "state", "annual",
                "reconciliation", "number_formats", "identity")

# The four state lines, in the order the handoff declares its accessors. The
# names are not written here: they are read from the contract and paired with
# the rows the manifest gives them.
STATE_KEYS = ("distribution_state", "profile_state", "profile_px", "year_count")


def build_phase8_inspection(spec: WorkbookSpec, calc: CalcContract,
                            raw_sim: dict[str, Any], row_window: int) -> dict[str, Any]:
    results = (spec.phase6_shell or {}).get("results")
    if not results:
        raise ValueError(
            "workbook.yaml carries no phase6_shell.results; there is no Results "
            "surface to project")
    if "annual" not in results or "reconciliation" not in results:
        raise ValueError(
            "workbook.yaml declares no Phase-8 annual or reconciliation block; the "
            "projection would describe a sheet that does not exist")

    annual = results["annual"]
    reconciliation = results["reconciliation"]
    handoff = raw_sim["sim_data"]["annual_records"]["handoff"]
    accessors = [str(entry["name"]) for entry in handoff["accessors"]]
    if len(accessors) != len(STATE_KEYS):
        raise ValueError(
            f"the handoff declares {len(accessors)} accessors; the Results state "
            f"block presents {len(STATE_KEYS)}")

    tolerances = calc.tolerances
    return {
        "schema_version": SCHEMA_VERSION,
        "purpose": (
            "Where the Phase-8 Results output surface sits. Addresses, procedure "
            "names and formats projected from workbook.yaml and the accepted "
            "contracts; no expected value and no tolerance of its own."
        ),
        "provenance": {
            "workbook_manifest": "workbook.yaml",
            "calc_contract_version": str(calc.version),
        },
        "sheet": str(results["sheet"]),
        "columns": {
            "label": str(results["label_column"]),
            "nominal": str(results["nominal_column"]),
            "pv": str(results["pv_column"]),
        },
        "run_stamp": {
            "first_row": int(results["run_stamp"]["first_row"]),
            "last_row": int(results["run_stamp"]["last_row"]),
            "fields": [{"key": str(f["key"]), "row": int(f["row"]),
                        "label": str(f["label"])}
                       for f in results["run_stamp"]["fields"]],
        },
        "summary": {
            "header_row": int(results["summary"]["header_row"]),
            "metrics": [{"key": str(m["key"]), "row": int(m["row"]),
                         "label": str(m["label"])}
                        for m in results["summary"]["metrics"]],
        },
        "selected": {
            "confidence_level_row": int(results["selected"]["confidence_level_row"]),
            # THE TWO LADDERS, KEPT APART IN THE PROJECTION TOO. A runner that
            # had to guess which row was the total would guess the way W5 did.
            "total_row": int(results["selected"]["quantile_row"]),
            "contingency_row": int(results["selected"]["contingency_row"]),
        },
        "state": {
            key: {
                "row": int(annual[f"{key}_row"]),
                "label": str(annual["labels"][key]),
                # The Phase-8 adapter, paired with the Phase-7 accessor whose
                # answer it returns. Both names, because a runner has to be able
                # to say which one it is really testing.
                "accessor": accessor,
                "procedure": "PCCM_Results" + accessor[len("PCCM_"):],
            }
            for key, accessor in zip(STATE_KEYS, accessors)
        },
        "annual": {
            "heading_row": int(annual["heading_row"]),
            "header_row": int(annual["header_row"]),
            "first_row": int(annual["first_row"]),
            # THE STRUCTURAL MAXIMUM, so a runner reads past the answer to prove
            # the window blanks rather than reading exactly as far as it hopes.
            "row_window": int(row_window),
            "columns": [{"key": str(c["key"]), "header": str(c["header"]),
                         "column": str(c["column"]), "source": str(c["source"]),
                         "field": str(c["field"]), "format": str(c["format"])}
                        for c in annual["columns"]],
        },
        "reconciliation": {
            "heading_row": int(reconciliation["heading_row"]),
            "header_row": int(reconciliation["header_row"]),
            "rows": {str(entry["key"]): int(reconciliation["first_row"]) + index
                     for index, entry in enumerate(reconciliation["rows"])},
            "verdicts": {str(k): str(v) for k, v in reconciliation["verdicts"].items()},
        },
        "number_formats": {str(k): str(v) for k, v in results["number_formats"].items()},
        # THE PROJECT'S OWN IDENTITY RULE, carried rather than restated. A runner
        # that typed 1e-12 would be a second tolerance authority the moment the
        # contract moved.
        "identity": {
            "identity_absolute_floor": float(tolerances.identity_absolute_floor),
            "identity_relative_coefficient": float(
                tolerances.identity_relative_coefficient),
            "conditioning_scale_floor": float(tolerances.conditioning_scale_floor),
            "identity_rule": (
                "|delta| <= max(identity_absolute_floor, identity_relative_coefficient "
                "* max(conditioning_scale_floor, conditioning_scale)); the conditioning "
                "scale names the magnitude of the arithmetic performed, never the "
                "magnitude of its net result."
            ),
        },
    }


def validate_phase8_inspection(inspection: dict[str, Any]) -> None:
    unexpected = sorted(set(inspection) - set(ALLOWED_KEYS))
    if unexpected:
        raise ValueError(f"{INSPECTION_FILENAME}: unexpected key(s) {unexpected}")

    # THE TWO LADDERS MAY NEVER COLLAPSE INTO ONE ROW. W5 read the contingency
    # block as the total and failed four reconciliations by exactly the
    # deterministic base; a projection that gave both the same row would let a
    # runner make that mistake again with the projection's blessing.
    selected = inspection["selected"]
    if selected["total_row"] == selected["contingency_row"]:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the total and contingency rows are the same row; "
            "they are different quantities")

    state = inspection["state"]
    if sorted(state) != sorted(STATE_KEYS):
        raise ValueError(f"{INSPECTION_FILENAME}: the state block presents {sorted(state)}")
    for key, entry in state.items():
        if not entry["procedure"].endswith(entry["accessor"][len("PCCM_"):]):
            raise ValueError(
                f"{INSPECTION_FILENAME}: the {key} adapter {entry['procedure']!r} does "
                f"not wrap the accessor {entry['accessor']!r}")
        if entry["procedure"] == entry["accessor"]:
            raise ValueError(
                f"{INSPECTION_FILENAME}: the {key} adapter is the accessor itself; a "
                "zero-argument accessor in a cell would answer once and never again")

    annual = inspection["annual"]
    last = annual["first_row"] + annual["row_window"] - 1
    if last >= inspection["reconciliation"]["heading_row"]:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the annual window reaches row {last}, at or past "
            f"the reconciliation heading")
    sources = {column["source"] for column in annual["columns"]}
    if not sources <= {"index", "profile"}:
        raise ValueError(f"{INSPECTION_FILENAME}: unknown annual record source(s) {sources}")
    if inspection["identity"]["identity_absolute_floor"] <= 0:
        raise ValueError(f"{INSPECTION_FILENAME}: the identity floor must be positive")


def emit_phase8_results(spec: WorkbookSpec, calc: CalcContract,
                        raw_sim: dict[str, Any], row_window: int,
                        build_dir: Path) -> Path:
    inspection = build_phase8_inspection(spec, calc, raw_sim, row_window)
    validate_phase8_inspection(inspection)
    path = build_dir / INSPECTION_FILENAME
    write_lf_artifact(path, json.dumps(inspection, indent=2) + "\n")
    return path
