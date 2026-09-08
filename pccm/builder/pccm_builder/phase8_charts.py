#!/usr/bin/env python3
"""The Phase-8 chart projection: WHAT is plotted, FROM WHERE, and WHEN NOT.

WHY A FIFTH PROJECTION AND NOT A WIDER FOURTH. `phase8_results_inspection.json`
and `phase8_dashboard_inspection.json` are the artefacts the accepted P8-1 and
P8-2 Windows runs were produced against; widening either to carry chart geometry
would change a file whose identity IS that evidence. So the charts get their
own, on exactly the terms each earlier step got its own.

WHAT IT CARRIES. For every chart: its key, its type, where it is anchored, its
title, every series range and the category range, the bridge block each comes
from, the ultimate authority behind that block, and the cell whose state word
says whether the plot means anything. A test - or a later Windows runner - can
answer all of that from here without spelling a single address.

AND WHAT THE BRIDGE IS. Each block is labelled with its `kind`:

  derived_chart_only   computed here for a chart and read by nothing else
  mirrored             a cell copied from an authoritative surface, unchanged

Nothing in this file is an analytical output, an authority, or a tolerance.
"""

from __future__ import annotations

import json
from pathlib import Path
import re
from typing import Any

from .artifact_io import write_lf_artifact
from .workbook_builder import (
    TORNADO_ELIGIBILITY_FIELD as _TORNADO_ELIGIBILITY_FIELD,
)
from .spec_loader import WorkbookSpec

SCHEMA_VERSION = 1

INSPECTION_FILENAME = "phase8_charts_inspection.json"

ALLOWED_KEYS = ("schema_version", "purpose", "provenance", "bridge_sheet",
                "chart_sheet", "sensitivity_sheet", "sensitivity_endpoint",
                "sensitivity_source", "zero_variance_status", "bridge", "charts",
                "number_formats")

# THE CHART TYPES THIS PROJECT PERMITS. Two dimensions, three shapes. A third
# dimension carries no data here and distorts the comparison a chart exists to
# make, so it is refused by name rather than by taste.
CHART_KINDS = ("line", "column", "bar")

# WHERE EACH BRIDGE BLOCK ULTIMATELY COMES FROM. Named so a reader never has to
# trace a formula to find out whether a chart is plotting a published answer or
# something this layer invented.
AUTHORITIES = {
    "annual": "Results annual cash flow (the persisted selected-Px profile)",
    "distribution": ("Results summary minimum/maximum and the published "
                     "iteration column of the active bank"),
    "drivers": "Sensitivity ranked drivers (Phase-7 Spearman ranking)",
    "status": "Sensitivity availability line",
}


def _bridge_blocks(charts: dict[str, Any], window: int) -> dict[str, Any]:
    """The four blocks, plus the two rows that mark where the bridge begins."""
    bridge = charts["bridge"]
    sheet = str(charts["bridge_sheet"])
    counts = {
        "annual": window,
        "distribution": int(bridge["distribution"]["bin_count"]),
        "drivers": int(bridge["drivers"]["top_n"]),
    }
    out: dict[str, Any] = {}
    for name in ("annual", "distribution", "drivers"):
        block = bridge[name]
        first = int(block["first_row"])
        out[name] = {
            "sheet": sheet,
            # DERIVED CHART-ONLY DATA, every one of them: computed for a chart,
            # read by nothing else, and an authority for nothing.
            "kind": "derived_chart_only",
            "authority": AUTHORITIES[name],
            "heading_row": int(block["heading_row"]),
            "header_row": int(block["header_row"]),
            "first_row": first,
            "last_row": first + counts[name] - 1,
            "row_count": counts[name],
            "columns": [{"key": str(column["key"]), "header": str(column["header"]),
                         "column": str(column["column"]),
                         "format": str(column["format"]),
                         "range": (f"{sheet}!${column['column']}${first}"
                                   f":${column['column']}${first + counts[name] - 1}")}
                        for column in block["columns"]],
        }
    status = bridge["status"]
    out["status"] = {
        "sheet": sheet,
        # MIRRORED, not derived: the sentence and every arm of it belong to the
        # Sensitivity sheet, and this block copies it so the Dashboard keeps
        # mirroring one surface.
        "kind": "mirrored",
        "authority": AUTHORITIES["status"],
        "heading_row": int(status["heading_row"]),
        "first_row": int(status["first_row"]),
        "rows": [{"key": str(entry["key"]), "label": str(entry["label"]),
                  "row": int(status["first_row"]) + index}
                 for index, entry in enumerate(status["rows"])],
    }
    return out


def build_phase8_charts_inspection(spec: WorkbookSpec, window: int,
                                   sim: Any) -> dict[str, Any]:
    shell = spec.phase6_shell or {}
    charts = shell.get("charts")
    if not charts:
        raise ValueError(
            "workbook.yaml carries no phase6_shell.charts; there is no chart layer "
            "to project")

    blocks = _bridge_blocks(charts, window)
    bridge = charts["bridge"]
    # WHERE THE WHOLE BRIDGE STARTS. A reader - or a control asking what lives
    # above the chart layer and what belongs to it - needs the boundary, not
    # only the four blocks inside it.
    blocks["heading_row"] = int(bridge["heading_row"])
    blocks["note_row"] = int(bridge["note_row"])
    projected = []
    for chart in charts["charts"]:
        source = str(chart["source"])
        block = blocks[source]
        columns = {column["key"]: column for column in block["columns"]}
        projected.append({
            "key": str(chart["key"]),
            "kind": str(chart["kind"]),
            "anchor": str(chart["anchor"]),
            "width_cm": float(chart["width"]),
            "height_cm": float(chart["height"]),
            "title": str(chart["title"]),
            "source_block": source,
            "authority": block["authority"],
            "categories": {
                "key": str(chart["categories"]),
                "range": columns[str(chart["categories"])]["range"],
            },
            "series": [{"key": str(series["key"]), "name": str(series["name"]),
                        "range": columns[str(series["key"])]["range"]}
                       for series in chart["series"]],
            # THE CELL WHOSE WORD SAYS WHETHER THE PLOT MEANS ANYTHING. A chart
            # carries no state logic of its own; it inherits the state its
            # source already publishes, and this names where a reader finds it.
            "state_source": str(chart["state_source"]),
            # A SECOND CONDITION WHERE ONE ANSWER IS NOT ENOUGH. The tornado
            # needs BOTH "does this ranked table belong to the published run?"
            # and "does that published run still match the model?": the first
            # compares two persisted records and is blind to a model that has
            # moved since, so either alone would have said CURRENT after the
            # request drifted. They do not collapse.
            "also_qualified_by": (str(chart["also_qualified_by"])
                                  if chart.get("also_qualified_by") else None),
            # WHAT THE CHART DOES WHEN THERE IS NOTHING TO PLOT. Not a promise
            # about pixels: a statement about the value the bridge supplies, and
            # NA() is the value every chart type declines to draw.
            "no_data_value": "NA()",
            "legend": len(chart["series"]) > 1,
        })

    # THE PUBLISHED SENSITIVITY SURFACE, PROJECTED BY KEY. The tornado bridge
    # declares which driver fields it mirrors; those same keys are looked up in
    # the sensitivity presentation block, so the two cannot drift apart and a
    # neighbouring column - `abs_rho` sits beside `rho` - cannot be picked up by
    # position.
    sensitivity = shell.get("sensitivity")
    if not sensitivity:
        raise ValueError(
            "workbook.yaml carries no phase6_shell.sensitivity; the tornado has "
            "no published ranking to mirror")
    declared = {str(column["key"]): str(column["column"])
                for column in sensitivity["columns"]}
    source_columns = []
    for column in bridge["drivers"]["columns"]:
        key = str(column["key"])
        if key not in declared:
            raise ValueError(
                f"{INSPECTION_FILENAME}: the tornado mirrors {key!r}, which the "
                "Sensitivity sheet does not publish")
        source_columns.append({"key": key, "column": declared[key]})
    # THE FIELD THAT DECIDES WHETHER A PUBLISHED ROW IS IN THE TORNADO AT ALL.
    # The sheet holds the ranked rows followed by the diagnostic ones, so a
    # consumer checking the chart against the sheet has to know which is which -
    # and must not decide for itself. Carried by the same name the bridge gates
    # on, so the two cannot disagree.
    eligibility = str(_TORNADO_ELIGIBILITY_FIELD)
    if eligibility not in declared:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the Sensitivity sheet publishes no "
            f"{eligibility!r} column; the tornado cannot tell a ranked driver "
            "from a diagnostic row")
    # THE ONE STATUS A CONSUMER OF THIS PROJECTION HAS TO RECOGNISE, carried
    # from the sensitivity contract that declares it. A runner proving a
    # zero-variance driver is absent from the chart must first find that driver
    # on the sheet, and typing the label would be a second declaration of a
    # contract string. It is NOT put in the Phase-6 gate-B cases, whose bytes are
    # digest-pinned Gate-B evidence: a Phase-8 need does not get to move those.
    zero_variance_status = str(
        sim.raw["sensitivity"]["zero_variance"]["status_label"])
    sensitivity_source = {
        "first_row": int(sensitivity["first_row"]),
        "row_window": int(sensitivity["row_window"]),
        "eligibility": {"key": eligibility, "column": declared[eligibility]},
        "columns": source_columns,
    }

    return {
        "schema_version": SCHEMA_VERSION,
        "purpose": (
            "The Phase-8 chart layer: which charts exist, their type, position, "
            "title, series and category ranges, the bridge block each reads and "
            "the authority behind it. Addresses projected from workbook.yaml; no "
            "expected value, no statistic and no state word of its own."
        ),
        "provenance": {"workbook_manifest": "workbook.yaml"},
        "bridge_sheet": str(charts["bridge_sheet"]),
        "chart_sheet": str(charts["chart_sheet"]),
        "sensitivity_sheet": str(charts["sensitivity_sheet"]),
        "sensitivity_endpoint": str(charts["sensitivity_endpoint"]),
        # WHERE THE RANKING THE TORNADO MIRRORS ACTUALLY SITS. The bridge block
        # above says where the ten plotted rows live on Results; this says which
        # published Sensitivity rows they must reproduce, so a consumer can
        # compare the two WITHOUT re-ranking anything.
        #
        # ONE FACT, ONE OWNER. Every value here is read off
        # phase6_shell.sensitivity - the presentation authority for that sheet -
        # and only the two fields the tornado mirrors are carried. The sheet's
        # NAME is not repeated: `sensitivity_sheet` above already declares it.
        "sensitivity_source": sensitivity_source,
        "zero_variance_status": zero_variance_status,
        "bridge": blocks,
        "charts": projected,
        "number_formats": {str(k): str(v)
                           for k, v in charts["number_formats"].items()},
        # THE HISTOGRAM'S BINNING RULE, CARRIED RATHER THAN RESTATED. A runner
        # that typed 20 would be a second authority the moment the manifest
        # moved, and a reader checking count conservation needs the edge rule.
        **{"bin_contract": {
            "bin_count": int(bridge["distribution"]["bin_count"]),
            "measure": str(bridge["distribution"]["measure"]),
            "edges": ("equal width across the published minimum and maximum; bin i "
                      "covers [lower, upper) except the last, which closes at "
                      "<= upper so the maximum is counted"),
            "empty": "no publication -> every bin is NA(), never 0",
            "degenerate": ("maximum <= minimum -> the first bin takes the whole "
                           "iteration count and the rest are NA()"),
        }},
    }


def validate_phase8_charts_inspection(inspection: dict[str, Any]) -> None:
    unexpected = sorted(set(inspection) - set(ALLOWED_KEYS) - {"bin_contract"})
    if unexpected:
        raise ValueError(f"{INSPECTION_FILENAME}: unexpected key(s) {unexpected}")

    if inspection["bridge_sheet"] == inspection["chart_sheet"]:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the bridge and the charts are on the same sheet; "
            "the Dashboard holds charts and the bridge lives where the data does")

    anchors: set[str] = set()
    keys: set[str] = set()
    for chart in inspection["charts"]:
        key = chart["key"]
        if key in keys:
            raise ValueError(f"{INSPECTION_FILENAME}: two charts are called {key!r}")
        keys.add(key)
        if chart["kind"] not in CHART_KINDS:
            raise ValueError(
                f"{INSPECTION_FILENAME}: chart {key!r} is a {chart['kind']!r} chart; "
                f"only {CHART_KINDS} are permitted")
        if chart["anchor"] in anchors:
            raise ValueError(
                f"{INSPECTION_FILENAME}: two charts are anchored at {chart['anchor']}")
        anchors.add(chart["anchor"])
        if not chart["series"]:
            raise ValueError(f"{INSPECTION_FILENAME}: chart {key!r} plots nothing")
        # EVERY RANGE NAMES THE BRIDGE SHEET. A range without one resolves
        # against the chart's own sheet, which holds none of this data.
        ranges = [chart["categories"]["range"]] + [s["range"] for s in chart["series"]]
        for reference in ranges:
            if not reference.startswith(inspection["bridge_sheet"] + "!"):
                raise ValueError(
                    f"{INSPECTION_FILENAME}: chart {key!r} reads {reference}, which is "
                    f"not on the bridge sheet")
        # AND THE CATEGORY IS NOT ALSO A SERIES. A chart plotting its own axis
        # against itself is a straight line that means nothing.
        if chart["categories"]["key"] in {s["key"] for s in chart["series"]}:
            raise ValueError(
                f"{INSPECTION_FILENAME}: chart {key!r} plots its category column as a "
                "series")
        if chart["legend"] != (len(chart["series"]) > 1):
            raise ValueError(
                f"{INSPECTION_FILENAME}: chart {key!r} declares a legend that does not "
                "match its series count")

    for name, block in inspection["bridge"].items():
        if not isinstance(block, dict):
            continue  # the two boundary rows, not a block
        if block["kind"] not in ("derived_chart_only", "mirrored"):
            raise ValueError(
                f"{INSPECTION_FILENAME}: bridge block {name!r} is an unlabelled kind "
                f"{block['kind']!r}")
        if not block["authority"]:
            raise ValueError(
                f"{INSPECTION_FILENAME}: bridge block {name!r} names no authority")

    # THE SOURCE THE TORNADO MIRRORS. A consumer dereferences every one of these
    # to read a published Sensitivity row; a missing field is not a failed
    # assertion on Windows, it is a terminated session, so it fails here.
    if not str(inspection.get("zero_variance_status", "")).strip():
        raise ValueError(
            f"{INSPECTION_FILENAME}: no zero_variance_status; a consumer could "
            "not tell a diagnostic row from a ranked one on the sheet")
    if "sensitivity_source" not in inspection:
        raise ValueError(
            f"{INSPECTION_FILENAME}: no sensitivity_source; the tornado has no "
            "projected route to the ranking it mirrors")
    source = inspection["sensitivity_source"]
    # THE ELIGIBILITY FIELD IS NOT ONE OF THE PLOTTED COLUMNS. It decides which
    # rows are in the chart at all; a source that carried it as a plotted field
    # would draw the rank.
    if "eligibility" not in source:
        raise ValueError(
            f"{INSPECTION_FILENAME}: sensitivity_source names no eligibility "
            "field; a diagnostic row could not be told from a ranked one")
    for field in ("first_row", "row_window", "columns"):
        if field not in source:
            raise ValueError(
                f"{INSPECTION_FILENAME}: sensitivity_source carries no {field!r}")
    eligibility = source["eligibility"]
    if not re.fullmatch(r"[A-Z]{1,3}", str(eligibility["column"])):
        raise ValueError(
            f"{INSPECTION_FILENAME}: the eligibility column "
            f"{eligibility['column']!r} is not a column letter")
    if any(str(column["key"]) == str(eligibility["key"])
           for column in source["columns"]):
        raise ValueError(
            f"{INSPECTION_FILENAME}: {eligibility['key']!r} is both the "
            "eligibility field and a plotted series")
    if int(source["first_row"]) < 1:
        raise ValueError(
            f"{INSPECTION_FILENAME}: sensitivity_source.first_row is not a row")
    # AND IT MIRRORS EXACTLY WHAT THE BRIDGE PLOTS - no more, no fewer, same
    # order. A source column the bridge does not carry would be a field nothing
    # plots; a bridge column with no source would be a plotted value with
    # nothing to check it against.
    plotted = [str(column["key"]) for column in inspection["bridge"]["drivers"]["columns"]]
    mirrored = [str(column["key"]) for column in source["columns"]]
    if mirrored != plotted:
        raise ValueError(
            f"{INSPECTION_FILENAME}: the tornado plots {plotted} and mirrors "
            f"{mirrored}; they must be the same fields in the same order")
    for column in source["columns"]:
        if not re.fullmatch(r"[A-Z]{1,3}", str(column["column"])):
            raise ValueError(
                f"{INSPECTION_FILENAME}: sensitivity_source column "
                f"{column['column']!r} is not a column letter")
    # A MAGNITUDE IS NOT A SIGNED CORRELATION. The tornado shows direction; a
    # source pointed at the absolute column would lose every negative driver's
    # sign and no plotted value would look wrong.
    if any(str(column["key"]).startswith("abs_") for column in source["columns"]):
        raise ValueError(
            f"{INSPECTION_FILENAME}: the tornado mirrors an absolute magnitude; "
            "the sign is what the chart exists to show")

    # THE ENDPOINT IS AN ENDPOINT, not a worksheet function. A chart that could
    # be produced by a cell would be a chart that reran an analysis to draw
    # itself.
    if not str(inspection["sensitivity_endpoint"]).startswith("PCCM_Run"):
        raise ValueError(
            f"{INSPECTION_FILENAME}: {inspection['sensitivity_endpoint']!r} is not a "
            "run endpoint")

    contract = inspection["bin_contract"]
    if contract["bin_count"] < 2:
        raise ValueError(
            f"{INSPECTION_FILENAME}: {contract['bin_count']} bins is not a histogram")


def emit_phase8_charts(spec: WorkbookSpec, window: int, build_dir: Path,
                       sim: Any) -> Path:
    inspection = build_phase8_charts_inspection(spec, window, sim)
    validate_phase8_charts_inspection(inspection)
    path = build_dir / INSPECTION_FILENAME
    write_lf_artifact(path, json.dumps(inspection, indent=2) + "\n")
    return path
