#!/usr/bin/env python3
"""The Phase-8 Dashboard projection: WHERE the P8-2 executive summary lives.

WHY A FOURTH PROJECTION AND NOT A WIDER THIRD. `phase8_results_inspection.json`
is the artefact the accepted P8-1 Windows run was produced against; widening it
to carry a Dashboard layout would change a file whose identity IS that evidence.
So the Dashboard gets its own, on exactly the terms Phase 7 and P8-1 each got
their own.

WHAT IT CARRIES THAT THE OTHERS DO NOT: the MAPPING. Every Dashboard row is
recorded with the Results block, key, row and column it mirrors, so a reader -
a test, or a later Windows runner - can check the sheet against the cell it
claims to copy rather than against a number typed into a harness. That mapping
is the whole claim of this step: the Dashboard owns nothing, and this file is
where that is checkable.

IDENTITIES ONLY. No expected value, no tolerance, no state word and no verdict:
those belong to Results, which is the point.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .artifact_io import write_lf_artifact
from .spec_loader import WorkbookSpec

SCHEMA_VERSION = 1

INSPECTION_FILENAME = "phase8_dashboard_inspection.json"

ALLOWED_KEYS = ("schema_version", "purpose", "provenance", "sheet", "source_sheet",
                "columns", "source_columns", "mirror_formula", "sections",
                "chart_region", "number_formats")

# The Results blocks a Dashboard row may mirror. Named here so a fifth block
# cannot be mirrored without this file being edited to say so.
SOURCE_BLOCKS = ("run_stamp", "summary", "selected", "state", "reconciliation",
                 # P8-3. The chart bridge's status rows live ON RESULTS, so a
                 # Dashboard mirror of one is still a mirror of Results and the
                 # accepted P8-2 rule - this sheet reads one surface - holds.
                 "chart_status")


def _results_rows(results: dict[str, Any]) -> dict[str, dict[str, int]]:
    """The same block-and-key to row resolution the renderer performs, from the
    same manifest structure. Duplicated nowhere: the renderer and this file both
    read `results`, and a test asserts the two agree cell for cell."""
    annual = results["annual"]
    selected = results["selected"]
    reconciliation = results["reconciliation"]
    return {
        "run_stamp": {str(f["key"]): int(f["row"])
                      for f in results["run_stamp"]["fields"]},
        "summary": {str(m["key"]): int(m["row"])
                    for m in results["summary"]["metrics"]},
        "selected": {
            "confidence_level": int(selected["confidence_level_row"]),
            "total": int(selected["quantile_row"]),
            "contingency": int(selected["contingency_row"]),
        },
        "state": {key: int(annual[f"{key}_row"])
                  for key in ("distribution_state", "profile_state",
                              "profile_px", "year_count")},
        "reconciliation": {str(e["key"]): int(reconciliation["first_row"]) + index
                           for index, e in enumerate(reconciliation["rows"])},
    }


def _results_labels(results: dict[str, Any]) -> dict[str, dict[str, str]]:
    return {
        "run_stamp": {str(f["key"]): str(f["label"])
                      for f in results["run_stamp"]["fields"]},
        "summary": {str(m["key"]): str(m["label"])
                    for m in results["summary"]["metrics"]},
        "selected": {
            "confidence_level": str(results["selected"]["labels"]["confidence_level"]),
            "total": str(results["selected"]["labels"]["quantile"]),
            "contingency": str(results["selected"]["labels"]["contingency"]),
        },
        "state": {str(k): str(v) for k, v in results["annual"]["labels"].items()},
        "reconciliation": {str(e["key"]): str(e["label"])
                           for e in results["reconciliation"]["rows"]},
    }


def _chart_status(charts: dict[str, Any] | None) -> tuple[dict[str, int], dict[str, str]]:
    if not charts:
        return {}, {}
    status = charts["bridge"]["status"]
    first = int(status["first_row"])
    rows = {str(e["key"]): first + index for index, e in enumerate(status["rows"])}
    labels = {str(e["key"]): str(e["label"]) for e in status["rows"]}
    return rows, labels


def build_phase8_dashboard_inspection(spec: WorkbookSpec) -> dict[str, Any]:
    shell = spec.phase6_shell or {}
    dashboard = shell.get("dashboard")
    results = shell.get("results")
    if not dashboard:
        raise ValueError(
            "workbook.yaml carries no phase6_shell.dashboard; there is no Dashboard "
            "surface to project")
    if not results:
        raise ValueError(
            "workbook.yaml carries no phase6_shell.results; the Dashboard mirrors "
            "Results and cannot be projected without it")

    rows = _results_rows(results)
    labels = _results_labels(results)
    status_rows, status_labels = _chart_status(shell.get("charts"))
    if status_rows:
        rows["chart_status"] = status_rows
        labels["chart_status"] = status_labels
    source_columns = {
        "nominal": str(results["nominal_column"]),
        "pv": str(results["pv_column"]),
    }
    nominal_col = str(dashboard["nominal_column"])
    pv_col = str(dashboard["pv_column"])
    source_sheet = str(dashboard["source_sheet"])

    sections = []
    for section in dashboard["sections"]:
        paired = bool(section.get("headers"))
        measures = ("nominal", "pv") if paired else ("nominal",)
        entries = []
        first = int(section["first_row"])
        for offset, entry in enumerate(section["rows"]):
            block = str(entry["source_block"])
            key = str(entry["source_key"])
            row = first + offset
            entries.append({
                "key": str(entry["key"]),
                "row": row,
                # THE LABEL RESULTS SHOWS, carried rather than re-typed. Which
                # rung `quantile_10` spells belongs to the simulation contract.
                "label": labels[block][key],
                "source_block": block,
                "source_key": key,
                "source_row": rows[block][key],
                "format": str(entry["format"]),
                # THE MAPPING, SPELLED OUT BOTH WAYS: which Dashboard cell, and
                # which Results cell it must equal. A reader checking this sheet
                # never has to work out an address.
                "cells": {
                    measure: {
                        "dashboard": f"{nominal_col if measure == 'nominal' else pv_col}{row}",
                        "results": (f"{source_sheet}!"
                                    f"{source_columns[measure]}{rows[block][key]}"),
                    }
                    for measure in measures
                },
            })
        sections.append({
            "key": str(section["key"]),
            "title": str(section["title"]),
            "heading_row": int(section["row"]),
            "note_row": int(section["note_row"]),
            "header_row": int(section["header_row"]) if paired else None,
            "first_row": first,
            "measures": list(measures),
            "rows": entries,
        })

    region = dashboard["chart_region"]
    return {
        "schema_version": SCHEMA_VERSION,
        "purpose": (
            "Where the Phase-8 Dashboard executive summary sits, and which Results "
            "cell each of its cells mirrors. Addresses and formats projected from "
            "workbook.yaml; no expected value, no state word and no verdict - those "
            "belong to Results."
        ),
        "provenance": {"workbook_manifest": "workbook.yaml"},
        "sheet": str(dashboard["sheet"]),
        "source_sheet": source_sheet,
        "columns": {
            "label": str(dashboard["label_column"]),
            "nominal": nominal_col,
            "pv": pv_col,
        },
        "source_columns": source_columns,
        "mirror_formula": str(dashboard["mirror_formula"]),
        "sections": sections,
        # RESERVED AND EMPTY, and recorded so a test can prove both: that the
        # summary stops above it, and that nothing is drawn inside it yet.
        "chart_region": {
            "heading_row": int(region["heading_row"]),
            "note_row": int(region["note_row"]),
            "first_row": int(region["first_row"]),
            "last_row": int(region["last_row"]),
        },
        "number_formats": {str(k): str(v)
                           for k, v in dashboard["number_formats"].items()},
    }


def validate_phase8_dashboard_inspection(inspection: dict[str, Any]) -> None:
    unexpected = sorted(set(inspection) - set(ALLOWED_KEYS))
    if unexpected:
        raise ValueError(f"{INSPECTION_FILENAME}: unexpected key(s) {unexpected}")

    # THE DASHBOARD MAY NOT BE ITS OWN SOURCE. A sheet that mirrored itself
    # would satisfy every mapping check in this file and mean nothing.
    if inspection["sheet"] == inspection["source_sheet"]:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the Dashboard mirrors itself; Results is the "
            "authority, not this sheet")

    template = inspection["mirror_formula"]
    if template.count("{ref}") != 2 or '=""' not in template:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the mirror does not guard against a blank source; "
            "Excel reads an empty reference back as 0 and the sheet would print a "
            "fabricated zero where no result exists")

    seen_rows: dict[int, str] = {}
    seen_dashboard_cells: set[str] = set()
    region = inspection["chart_region"]
    for section in inspection["sections"]:
        for entry in section["rows"]:
            if entry["source_block"] not in SOURCE_BLOCKS:
                raise ValueError(
                    f"{INSPECTION_FILENAME}: {entry['key']!r} mirrors unknown Results "
                    f"block {entry['source_block']!r}")
            if entry["row"] in seen_rows:
                raise ValueError(
                    f"{INSPECTION_FILENAME}: rows {seen_rows[entry['row']]!r} and "
                    f"{entry['key']!r} both occupy row {entry['row']}")
            seen_rows[entry["row"]] = entry["key"]
            if entry["row"] >= region["heading_row"]:
                raise ValueError(
                    f"{INSPECTION_FILENAME}: {entry['key']!r} sits at row "
                    f"{entry['row']}, inside the reserved chart region")
            for measure, cells in entry["cells"].items():
                if cells["dashboard"] in seen_dashboard_cells:
                    raise ValueError(
                        f"{INSPECTION_FILENAME}: {cells['dashboard']} is written twice")
                seen_dashboard_cells.add(cells["dashboard"])
                if not cells["results"].startswith(inspection["source_sheet"] + "!"):
                    raise ValueError(
                        f"{INSPECTION_FILENAME}: {entry['key']!r} {measure} reads "
                        f"{cells['results']}, which is not on the source sheet")

    # NO RESULTS CELL IS MIRRORED TWICE, and this is the rule that matters most
    # on this sheet.
    #
    # IT WAS NOT THE FIRST RULE WRITTEN HERE. The first pair of checks named the
    # two hazards directly - total against contingency, selected level against
    # profile Px - and a mutation walked through both of them, because
    # re-pointing the TOTAL at the contingency's key does not leave a "total"
    # entry for a check that looks one up by name to find. The general rule has
    # no such blind spot: whatever a row is called, if two rows read one Results
    # cell then the sheet is showing one quantity under two labels, and nothing
    # on it recomputes the difference that would expose the substitution.
    duplicated: dict[tuple[str, str], str] = {}
    for section in inspection["sections"]:
        for entry in section["rows"]:
            source = (entry["source_block"], entry["source_key"])
            if source in duplicated:
                raise ValueError(
                    f"{INSPECTION_FILENAME}: {duplicated[source]!r} and "
                    f"{entry['key']!r} both mirror {source[0]}.{source[1]}; one "
                    "Results cell shown under two labels is one quantity wearing "
                    "another's name")
            duplicated[source] = entry["key"]

    # THE TWO LADDERS MAY NEVER COLLAPSE INTO ONE ROW - the W5 defect, restated
    # for this sheet. Kept alongside the general rule because it names the
    # specific quantities, and a reader of this file should not have to derive
    # which two rows the project has already confused once.
    mirrored = {(e["source_block"], e["source_key"]): e["source_row"]
                for section in inspection["sections"] for e in section["rows"]}
    total = mirrored.get(("selected", "total"))
    contingency = mirrored.get(("selected", "contingency"))
    if total is not None and total == contingency:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the total and the contingency mirror the same "
            "Results row; they are different quantities")

    # AND THE SELECTED LEVEL IS NOT THE PUBLISHED PROFILE'S Px. One is what the
    # reader is asking for now; the other is what the stored profile was blended
    # at. P8-1 Part B exists because they can differ, and a Dashboard that read
    # one from the other's row would make that part unobservable here.
    selected_level = mirrored.get(("selected", "confidence_level"))
    profile_px = mirrored.get(("state", "profile_px"))
    if selected_level is not None and selected_level == profile_px:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the selected confidence level and the published "
            "profile Px mirror the same Results row; they are different questions")


def emit_phase8_dashboard(spec: WorkbookSpec, build_dir: Path) -> Path:
    inspection = build_phase8_dashboard_inspection(spec)
    validate_phase8_dashboard_inspection(inspection)
    path = build_dir / INSPECTION_FILENAME
    write_lf_artifact(path, json.dumps(inspection, indent=2) + "\n")
    return path
