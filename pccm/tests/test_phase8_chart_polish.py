#!/usr/bin/env python3
"""FINAL-DELIVERY CHART POLISH: the four Dashboard charts, presentation only.

WHAT WAS SEEN IN THE REAL WORKBOOK. The two year charts were bound to the whole
reserved annual window - 200 rows, the structural year ceiling - so a project of
ten years sat in the leftmost twentieth of its axis with its labels on top of
each other. The tornado's value axis was automatic and read three decimals. The
histogram printed all twenty bin-edge captions on an 18 cm axis.

WHAT CHANGED, AND WHERE. The manifest (workbook.yaml) declares an applied-year
binding on the annual bridge block, a fixed value-axis scale on the tornado, a
label interval on the histogram and the s-curve's new title; the builder
projects one sheet-scoped defined name per annual column, cut at the SAME year
count cell the bridge guard reads, and plots the year charts through those
names; the Phase-8 chart projection carries and validates every one of these.

WHAT DID NOT CHANGE. No bridge formula, no bin, no edge, no count, no ranking,
no rho, no published value. Every check below that says so compares against the
previous head by text or by value.
"""

from __future__ import annotations

import copy
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

import pytest  # noqa: E402
import yaml  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

from pccm_builder.phase8_charts import (  # noqa: E402
    build_phase8_charts_inspection,
    validate_phase8_charts_inspection,
)
from pccm_builder.spec_loader import SpecError, load_spec  # noqa: E402
from pccm_builder.structure_loader import load_structure_contract  # noqa: E402
from pccm_builder.workbook_builder import (  # noqa: E402
    _annual_extent_cell,
    applied_year_bindings,
)

SPEC = PCCM_ROOT / "spec"
BUILD = PCCM_ROOT / "build"
MANIFEST = SPEC / "workbook.yaml"
# THE HEAD BEFORE THIS CORRECTION: the runner tail audit, production at 548799f.
PREVIOUS_HEAD = "ebeae65"

YEAR_CHARTS = ("s_curve", "annual_cash_flow")
_CACHE: dict = {}


def _projection() -> dict:
    if "charts" not in _CACHE:
        _CACHE["charts"] = json.loads(
            (BUILD / "phase8_charts_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["charts"]


def _shell() -> dict:
    if "shell" not in _CACHE:
        _CACHE["shell"] = yaml.safe_load(MANIFEST.read_text(encoding="utf-8"))["phase6_shell"]
    return _CACHE["shell"]


def _workbook():
    if "workbook" not in _CACHE:
        manifest = json.loads((BUILD / "stage_b_manifest.json").read_text(encoding="utf-8"))
        _CACHE["workbook"] = load_workbook(BUILD / manifest["stage_a_filename"])
    return _CACHE["workbook"]


def _by_key() -> dict[str, dict]:
    return {chart["key"]: chart for chart in _projection()["charts"]}


def _chart_objects() -> dict[str, object]:
    return {chart.title.tx.rich.p[0].r[0].t: chart
            for chart in _workbook()[_projection()["chart_sheet"]]._charts}


def _chart_parts() -> list[str]:
    manifest = json.loads((BUILD / "stage_b_manifest.json").read_text(encoding="utf-8"))
    with zipfile.ZipFile(BUILD / manifest["stage_a_filename"]) as archive:
        return [archive.read(name).decode("utf-8") for name in archive.namelist()
                if re.fullmatch(r"xl/charts/chart\d+\.xml", name)]


def _part_for(title: str) -> str:
    for part in _chart_parts():
        if f"<a:t>{title}</a:t>" in part:
            return part
    raise AssertionError(f"no chart part carries the title {title!r}")


def _window() -> int:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    return int(structure.limits.max_generated_year_columns)


def _git_show(path: str) -> str:
    return subprocess.run(["git", "show", f"{PREVIOUS_HEAD}:pccm/{path}"],
                          cwd=PCCM_ROOT.parent, check=True,
                          stdout=subprocess.PIPE).stdout.decode("utf-8")


def _previous_charts() -> dict:
    if "previous" not in _CACHE:
        _CACHE["previous"] = yaml.safe_load(_git_show("spec/workbook.yaml"))["phase6_shell"]["charts"]
    return _CACHE["previous"]


# THE FORMULA EVERY APPLIED-YEAR NAME CARRIES, and the rows it selects for one
# stamped year count - what Excel does with INDEX, MAX, MIN and N, in Python.
_NAME_FORMULA = re.compile(
    r"^(?P<sheet>[A-Za-z_]+)!\$(?P<col>[A-Z]+)\$(?P<first>\d+):INDEX\("
    r"(?P=sheet)!\$(?P=col)\$(?P=first):\$(?P=col)\$(?P<last>\d+),"
    r"MAX\(1,MIN\((?P<window>\d+),N\((?P=sheet)!\$(?P<xcol>[A-Z]+)\$(?P<xrow>\d+)\)\)\)\)$")


def _rows_selected(formula: str, year_count) -> tuple[int, int]:
    match = _NAME_FORMULA.fullmatch(formula)
    assert match, formula
    first, last, window = int(match["first"]), int(match["last"]), int(match["window"])
    assert last - first + 1 == window, (first, last, window)
    # N(): a number is itself, text and blank are 0.
    n = float(year_count) if isinstance(year_count, (int, float)) else 0.0
    rows = max(1, min(window, int(n)))
    return first, first + rows - 1


# ===========================================================================
# 1. CUMULATIVE COST PROFILE
# ===========================================================================
def test_01_the_s_curve_is_titled_cumulative_cost_profile_and_is_still_a_line_chart() -> None:
    spec = _by_key()["s_curve"]
    assert spec["title"] == "Cumulative Cost Profile"
    assert spec["kind"] == "line"
    chart = _chart_objects()["Cumulative Cost Profile"]
    assert type(chart).__name__ == "LineChart"
    for series in chart.series:
        assert series.smooth is False or series.smooth is None
    # THE SAME SERIES AND CATEGORY AS BEFORE, by key and by name.
    was = {c["key"]: c for c in _previous_charts()["charts"]}["s_curve"]
    assert was["title"] == "Cumulative Cost"
    assert [(s["key"], s["name"]) for s in spec["series"]] == [
        (s["key"], s["name"]) for s in was["series"]]
    assert spec["categories"]["key"] == was["categories"] == "calendar_year"
    assert spec["source_block"] == was["source"] == "annual"
    assert spec["state_source"] == was["state_source"]


def test_02_the_year_charts_plot_the_applied_years_through_the_bound_names() -> None:
    """THE CORRECTION ITSELF. Categories and every value series of both year
    charts read the applied-year names, never the 200-row window - in the
    projection and in the chart parts the workbook actually carries."""
    sheet = _projection()["bridge_sheet"]
    block = _projection()["bridge"]["annual"]
    binding = block["applied_binding"]
    assert binding["name_prefix"] == "chartAnnual" and binding["extent"] == "year_count"
    assert binding["window_rows"] == _window() == block["row_count"]
    names = {column["key"]: column["applied_binding"] for column in block["columns"]}
    for key in YEAR_CHARTS:
        spec = _by_key()[key]
        chart = _chart_objects()[spec["title"]]
        for entry in [spec["categories"]] + spec["series"]:
            bound = names[entry["key"]]
            assert bound["name"] == f"chartAnnual_{entry['key']}"
            assert entry["binding"] == bound["name"]
            assert entry["range"] == f"{sheet}!{bound['name']}"
            assert entry["window_range"] == bound["window_range"]
            assert entry["range"] != entry["window_range"]
            # A CHART BOUND TO THE RESERVED TAIL WOULD END AT ROW 478.
            assert re.search(r"\$\d+:\$[A-Z]+\$\d+$", entry["range"]) is None
        assert [s.val.numRef.f for s in chart.series] == [
            f"'{sheet}'!chartAnnual_{s['key']}" for s in spec["series"]]
        for series in chart.series:
            assert series.cat.numRef.f == f"'{sheet}'!chartAnnual_calendar_year"
        part = _part_for(spec["title"])
        assert f"<f>'{sheet}'!chartAnnual_calendar_year</f>" in part
        assert re.search(r"<f>'" + sheet + r"'!\$[A-Z]+\$\d+:\$[A-Z]+\$\d+</f>", part) is None, (
            "a year chart still carries a cell-range reference")


def test_03_the_names_are_sheet_scoped_and_cut_at_the_one_year_count_cell() -> None:
    """ONE YEAR-COUNT AUTHORITY. Every name's formula is its column's window
    range cut at the Results annual state cell the bridge guard already reads;
    the guard's own formula at the first row names the same cell."""
    results = _workbook()[_projection()["bridge_sheet"]]
    block = _projection()["bridge"]["annual"]
    binding = block["applied_binding"]
    shell = _shell()
    expected_extent = ("Results!" + _annual_extent_cell(shell["results"], "year_count"))
    assert binding["extent_cell"] == expected_extent
    # THE GUARD READS THE SAME CELL: `1>$D$56` at the first bridge row.
    first = int(block["first_row"])
    guard_cell = expected_extent.split("!", 1)[1]
    for column in block["columns"]:
        formula = results[f"{column['column']}{first}"].value
        assert f",1>{guard_cell})" in formula, (column["key"], formula)
    # THE NAMES, ON THE BRIDGE SHEET, EXACTLY AS PROJECTED.
    on_sheet = {name: entry.attr_text for name, entry in results.defined_names.items()}
    assert set(on_sheet) == {column["applied_binding"]["name"] for column in block["columns"]}
    for column in block["columns"]:
        bound = column["applied_binding"]
        assert on_sheet[bound["name"]] == bound["formula"]
        assert bound["scope"] == _projection()["bridge_sheet"]
        assert bound["window_range"] == column["range"]
        assert bound["extent_cell"] == expected_extent
        assert _NAME_FORMULA.fullmatch(bound["formula"]), bound["formula"]
    # AND NOT WORKBOOK-LEVEL: the accepted contract-derived name set is untouched.
    assert not any(name.startswith("chartAnnual") for name in _workbook().defined_names)
    # THE BUILDER AND THE PROJECTION AGREE, name for name.
    live = applied_year_bindings(shell["charts"], shell["results"], _window())
    assert {k: v for k, v in live.items()} == {
        column["key"]: column["applied_binding"] for column in block["columns"]}


@pytest.mark.parametrize("year_count,expected_rows", [
    (1, 1), (3, 3), (7, 7), (10, 10), (30, 30), (75, 75), (200, 200),
    # OUTSIDE THE WINDOW OR NOT A COUNT: the names stay inside the block and
    # never below one row, whose value is NA() and whose category is blank.
    (250, 200), (0, 1), ("", 1), ("NOT PRODUCED", 1),
])
def test_04_every_applied_duration_selects_exactly_its_own_rows(year_count, expected_rows) -> None:
    """MULTIPLE DURATIONS, NOT THE DEMO'S. For each stamped year count the
    names select first_row .. first_row + n - 1 of the block, and nothing of
    the reserved tail beyond; the bridge guard turns every value past n to
    NA() and every category to blank, so the two agree row for row."""
    block = _projection()["bridge"]["annual"]
    first = int(block["first_row"])
    results = _workbook()[_projection()["bridge_sheet"]]
    guard_cell = block["applied_binding"]["extent_cell"].split("!", 1)[1]
    for column in block["columns"]:
        start, end = _rows_selected(column["applied_binding"]["formula"], year_count)
        assert start == first
        assert end - start + 1 == expected_rows, (column["key"], year_count)
        assert end <= int(block["last_row"])
        # THE FIRST ROW PAST THE SELECTION IS GUARDED OUT BY THE SAME COUNT.
        if end < int(block["last_row"]) and isinstance(year_count, int) and 0 < year_count < 200:
            beyond = results[f"{column['column']}{end + 1}"].value
            assert f",{year_count + 1}>{guard_cell})" in beyond, beyond
            assert (',"",' in beyond) if column["absent"] == '""' else (",NA()," in beyond)


def test_05_a_year_chart_bound_back_to_the_reserved_window_is_refused() -> None:
    """THE REGRESSION THIS CORRECTION FORBIDS: a year chart reading the whole
    200-row window again - as the projection, and as the manifest."""
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    from pccm_builder import load_sim_contract
    inspection = build_phase8_charts_inspection(
        load_spec(MANIFEST), structure.limits.max_generated_year_columns,
        load_sim_contract(SPEC / "sim_contract.yaml"))
    validate_phase8_charts_inspection(inspection)
    assert inspection == _projection()
    for key in YEAR_CHARTS:
        for field in ("categories", "series"):
            damaged = copy.deepcopy(inspection)
            chart = next(c for c in damaged["charts"] if c["key"] == key)
            entry = chart["categories"] if field == "categories" else chart["series"][0]
            entry["range"] = entry["window_range"]
            entry["binding"] = None
            with pytest.raises(ValueError, match="applied-year name|whole reserved window|whole window"):
                validate_phase8_charts_inspection(damaged)
        damaged = copy.deepcopy(inspection)
        chart = next(c for c in damaged["charts"] if c["key"] == key)
        chart["series"][0]["range"] = chart["series"][0]["window_range"]
        with pytest.raises(ValueError):
            validate_phase8_charts_inspection(damaged)
    # A NAME THAT DOES NOT CUT AT THE EXTENT CELL, OR OUTSIDE THE WINDOW.
    for mutate in (
        lambda b: b.__setitem__("formula", b["formula"].replace("MIN(200,", "MIN(250,")),
        lambda b: b.__setitem__("formula", b["formula"].replace("MAX(1,", "MAX(0,")),
        lambda b: b.__setitem__("formula", b["formula"].replace("$D$56", "$D$57")),
        lambda b: b.__setitem__("extent_cell", "Results!$D$57"),
    ):
        damaged = copy.deepcopy(inspection)
        mutate(damaged["bridge"]["annual"]["columns"][1]["applied_binding"])
        with pytest.raises(ValueError):
            validate_phase8_charts_inspection(damaged)
    # AND THE MANIFEST WITHOUT THE BINDING BUILDS WINDOW-BOUND CHARTS AGAIN,
    # which is exactly the shape the projection's own validator accepts as
    # "no binding declared" - so the binding's presence is pinned here.
    assert _shell()["charts"]["bridge"]["annual"]["applied_binding"] == {
        "name_prefix": "chartAnnual", "extent": "year_count"}


def test_06_the_manifest_refuses_a_binding_it_cannot_resolve() -> None:
    text = MANIFEST.read_text(encoding="utf-8")
    for old, new, message in (
        ('          extent: "year_count"\n', '          extent: "years_stamped"\n',
         "not a row the Results annual block publishes"),
        ('          name_prefix: "chartAnnual"\n', '          name_prefix: "9charts"\n',
         "cannot begin a defined name"),
        ("          major_unit: 0.10\n", "          major_unit: 0.50\n", "is not an axis"),
        ("        category_label_interval: 2\n", "        category_label_interval: 0\n",
         "is not a positive count"),
    ):
        assert text.count(old) == 1, old
        path = MANIFEST.with_name("workbook.polish-mutation.yaml")
        path.write_text(text.replace(old, new, 1), encoding="utf-8")
        try:
            with pytest.raises(SpecError, match=re.escape(message)):
                load_spec(path)
        finally:
            path.unlink(missing_ok=True)


# ===========================================================================
# 2. ANNUAL CASH FLOW
# ===========================================================================
def test_10_the_annual_cash_flow_is_still_a_column_chart_of_the_published_profile() -> None:
    spec = _by_key()["annual_cash_flow"]
    assert spec["title"] == "Annual Cash Flow" and spec["kind"] == "column"
    chart = _chart_objects()["Annual Cash Flow"]
    assert type(chart).__name__ == "BarChart" and chart.type == "col"
    was = {c["key"]: c for c in _previous_charts()["charts"]}["annual_cash_flow"]
    assert [(s["key"], s["name"]) for s in spec["series"]] == [
        (s["key"], s["name"]) for s in was["series"]] == [("annual_nominal", "Annual Nominal")]
    assert spec["categories"]["key"] == was["categories"] == "calendar_year"
    assert spec["categories"]["binding"] == "chartAnnual_calendar_year"
    assert spec["series"][0]["binding"] == "chartAnnual_annual_nominal"
    assert spec["value_axis_scale"] is None and spec["category_label_interval"] is None


# ===========================================================================
# 3. TOP DRIVERS BY RANK CORRELATION
# ===========================================================================
def test_20_the_tornado_value_axis_is_fixed_at_0_to_0_45_by_0_10_and_reads_two_decimals() -> None:
    """THE EQUIVALENT REPRESENTATION, STATED. openpyxl calls the VALUE axis
    `y_axis` for every chart type, and on a horizontal bar chart Excel draws
    that axis along the bottom - so the "x-axis" minimum, maximum and major
    unit a reader sees are `<c:valAx><c:scaling><c:min/><c:max/></c:scaling>`
    and `<c:majorUnit/>` in the chart part, which is what this asserts."""
    spec = _by_key()["tornado"]
    assert spec["kind"] == "bar"
    assert spec["value_axis_scale"] == {"min": 0.0, "max": 0.45, "major_unit": 0.1}
    assert spec["value_axis_format"] == "0.00"
    chart = _chart_objects()["Top Drivers by Rank Correlation"]
    assert type(chart).__name__ == "BarChart" and chart.type == "bar"
    assert chart.y_axis.scaling.min == 0.0
    assert chart.y_axis.scaling.max == 0.45
    assert chart.y_axis.majorUnit == 0.1
    assert chart.y_axis.numFmt.formatCode == "0.00"
    assert chart.y_axis.numFmt.sourceLinked is False
    part = _part_for("Top Drivers by Rank Correlation")
    value_axis = part[part.index("<valAx>"): part.index("</valAx>")]
    assert '<min val="0"' in value_axis and '<max val="0.45"' in value_axis, value_axis
    assert '<majorUnit val="0.1"' in value_axis, value_axis
    assert 'formatCode="0.00"' in value_axis, value_axis
    # NO OTHER CHART HAS A FIXED SCALE.
    for key, other in _by_key().items():
        if key != "tornado":
            assert other["value_axis_scale"] is None, key


def test_21_the_tornado_ranking_source_and_rho_values_are_untouched() -> None:
    spec = _by_key()["tornado"]
    was = {c["key"]: c for c in _previous_charts()["charts"]}["tornado"]
    assert spec["anchor"] == was["anchor"]
    assert spec["width_cm"] == float(was["width"]) and spec["height_cm"] == float(was["height"])
    assert spec["plot_area"] == {k: float(v) for k, v in was["plot_area"].items()}
    assert spec["state_source"] == was["state_source"]
    assert spec["also_qualified_by"] == was["also_qualified_by"]
    assert spec["category_axis_format"] == _projection()["number_formats"][was["category_axis_format"]]
    assert spec["source_block"] == was["source"] == "drivers"
    assert spec["categories"]["key"] == was["categories"] == "driver_name"
    assert [(s["key"], s["name"]) for s in spec["series"]] == [
        (s["key"], s["name"]) for s in was["series"]] == [("rho", "Rho")]
    assert spec["categories"]["range"] == spec["categories"]["window_range"]
    assert spec["series"][0]["range"] == spec["series"][0]["window_range"]
    assert spec["categories"]["binding"] is None and spec["series"][0]["binding"] is None
    previous = _previous_charts()["bridge"]["drivers"]
    assert _shell()["charts"]["bridge"]["drivers"] == previous, "the tornado's bridge moved"
    assert [c["key"] for c in _projection()["sensitivity_source"]["columns"]] == ["driver_name", "rho"]
    # THE PUBLISHED SENSITIVITY SURFACE THE TORNADO MIRRORS IS THE PREVIOUS HEAD'S.
    previous_shell = yaml.safe_load(_git_show("spec/workbook.yaml"))["phase6_shell"]
    assert previous_shell["sensitivity"] == _shell()["sensitivity"], "the Sensitivity surface moved"
    # THE BRIDGE CELLS STILL CARRY rho AT THREE DECIMALS; only the axis reads two.
    assert _projection()["number_formats"]["rho"] == "0.000"
    block = _projection()["bridge"]["drivers"]
    rho = next(c for c in block["columns"] if c["key"] == "rho")
    cell = _workbook()[_projection()["bridge_sheet"]][f"{rho['column']}{block['first_row']}"]
    assert cell.number_format == "0.000"


# ===========================================================================
# 4. TOTAL COST DISTRIBUTION
# ===========================================================================
def test_30_the_histogram_prints_every_second_caption_and_plots_every_bin() -> None:
    spec = _by_key()["histogram"]
    assert spec["category_label_interval"] == 2
    chart = _chart_objects()["Total Cost Distribution"]
    assert chart.x_axis.tickLblSkip == 2
    part = _part_for("Total Cost Distribution")
    category_axis = part[part.index("<catAx>"): part.index("</catAx>")]
    assert '<tickLblSkip val="2"' in category_axis, category_axis
    # EVERY BIN IS STILL PLOTTED: the ranges are the whole block, as before.
    assert spec["categories"]["range"] == spec["categories"]["window_range"] == "Results!$B$487:$B$506"
    assert spec["series"][0]["range"] == spec["series"][0]["window_range"] == "Results!$F$487:$F$506"
    assert [s.val.numRef.f for s in chart.series] == ["'Results'!$F$487:$F$506"]
    assert chart.series[0].cat.numRef.f == "'Results'!$B$487:$B$506"
    # NO OTHER CHART CARRIES AN INTERVAL: the year axes are left to Excel.
    for key, other in _by_key().items():
        if key != "histogram":
            assert other["category_label_interval"] is None, key
            assert _chart_objects()[other["title"]].x_axis.tickLblSkip is None, key


def test_31_the_histogram_bins_edges_and_counts_are_untouched() -> None:
    previous = _previous_charts()["bridge"]["distribution"]
    assert _shell()["charts"]["bridge"]["distribution"] == previous, "the histogram's bridge moved"
    assert _projection()["bin_contract"]["bin_count"] == 20
    assert _projection()["bridge"]["distribution"]["row_count"] == 20


# ===========================================================================
# 5. NOTHING ANALYTICAL MOVED
# ===========================================================================
def _function_text(source: str, name: str) -> str:
    start = source.index(f"\ndef {name}(")
    end = source.index("\n\n\n", start)
    return source[start:end]


def test_40_the_bridge_formula_writers_are_the_previous_heads_and_the_extent_cell_is_the_same() -> None:
    """THE VALUE-PRODUCING FORMULAS ARE BYTE-IDENTICAL to the previous head for
    the histogram and the tornado, and the annual writer differs by one line:
    the guard's year-count cell is now read through the helper the names use,
    and the helper returns exactly the address the old line spelled."""
    now = (PCCM_ROOT / "builder" / "pccm_builder" / "workbook_builder.py").read_text(encoding="utf-8")
    before = _git_show("builder/pccm_builder/workbook_builder.py")
    for name in ("_chart_bridge_distribution", "_chart_bridge_drivers", "_chart_bridge_status",
                 "_absent", "_chart_ranges"):
        assert _function_text(now, name) == _function_text(before, name), name
    old_line = '    years = f"${nominal_col}${annual[\'year_count_row\']}"\n'
    new_lines = ('    # THE ONE YEAR-COUNT AUTHORITY: the same cell the applied-year chart names\n'
                 '    # cut at (_annual_extent_cell), so a chart can never outreach its guard.\n'
                 '    years = _annual_extent_cell(results, "year_count")\n')
    annual_before = _function_text(before, "_chart_bridge_annual")
    annual_now = _function_text(now, "_chart_bridge_annual")
    assert annual_before.count(old_line) == 1 and annual_now.count(new_lines) == 1
    assert annual_now.replace(new_lines, old_line) == annual_before
    results = _shell()["results"]
    assert _annual_extent_cell(results, "year_count") == (
        f"${results['nominal_column']}${results['annual']['year_count_row']}")
    # AND THE WORKBOOK'S BRIDGE FORMULAS ARE THE ONES THE PREVIOUS HEAD WROTE.
    sheet = _workbook()[_projection()["bridge_sheet"]]
    annual = _projection()["bridge"]["annual"]
    assert sheet[f"H{annual['first_row']}"].value == '=IF(OR($D$53="NOT PRODUCED",1>$D$56),NA(),SUM($F$59:$F$59))'
    assert sheet[f"D{annual['first_row']}"].value == '=IF(OR($D$53="NOT PRODUCED",1>$D$56),"",$D$59)'


def test_42_the_polish_is_declared_by_exact_reversal_to_the_previous_head() -> None:
    """DECLARATION OVER PROHIBITION. Every historical control that pins the
    manifest or the builder to an accepted tree keeps its claim through
    chart_polish_declaration: taking the polish off the current file must
    reproduce the previous head byte for byte, the accepted side passes
    through unchanged, and a partly present layer is refused."""
    sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from chart_polish_declaration import (ACCEPTED_BEFORE_CHART_POLISH,
                                          DECLARED_CHART_POLISH_CHANGES, strip_chart_polish)
    assert ACCEPTED_BEFORE_CHART_POLISH == PREVIOUS_HEAD
    assert set(DECLARED_CHART_POLISH_CHANGES) == {
        "spec/workbook.yaml", "builder/pccm_builder/workbook_builder.py",
        "builder/pccm_builder/phase8_charts.py", "builder/pccm_builder/spec_loader.py"}
    for name in DECLARED_CHART_POLISH_CHANGES:
        now = (PCCM_ROOT / name).read_text(encoding="utf-8")
        then = _git_show(name)
        assert now != then, name
        assert strip_chart_polish(name, now) == then, name
        assert strip_chart_polish(name, then) == then, name
        damaged = now.replace("value_axis_scale", "value_axis_scal3", 1)
        with pytest.raises(AssertionError, match="partly present"):
            strip_chart_polish(name, damaged)


def test_41_no_production_source_or_contract_moved() -> None:
    for path in ("src", "spec/sim_contract.yaml", "spec/calc_contract.yaml",
                 "spec/structure_contract.yaml", "spec/input_contract.yaml",
                 "spec/driver_contract.yaml"):
        done = subprocess.run(["git", "diff", "--quiet", PREVIOUS_HEAD, "--", f"pccm/{path}"],
                              cwd=PCCM_ROOT.parent)
        assert done.returncode == 0, f"{path} changed since {PREVIOUS_HEAD}"


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
