#!/usr/bin/env python3
"""Phase 8, Step 3: the analytical chart layer, proved against its sources.

WHAT THIS SLICE IS. Four charts on the Dashboard - a cumulative cost curve, a
total-cost histogram, an annual cash flow and a driver tornado - each pointed at
a bridge block on Results, and each bridge block derived from a surface that was
already accepted. No engine, no statistic, no ranking and no VBA.

WHAT THESE CONTROLS EXIST TO CATCH. A chart fails silently in ways a table does
not. Excel plots a zero-length string as ZERO, so a series over a 200-row window
draws 196 fabricated years of a project costing nothing, and it looks like data.
A tornado re-sorted by signed rho looks tidier and is wrong. A histogram whose
last bin is half-open loses the maximum and nobody counts. So the controls below
check the SHAPE OF THE ABSENCE as hard as they check the presence: NA() where
there is no point, count conservation across the bins, the rank order taken and
not re-taken.

WHAT THEY CANNOT PROVE: there is no Excel here. Nothing below claims a formula
was evaluated, a chart was drawn, or a series rendered any particular way.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
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

SPEC = PCCM_ROOT / "spec"
BUILD = PCCM_ROOT / "build"
SRC = PCCM_ROOT / "src" / "vba"
MANIFEST = SPEC / "workbook.yaml"

# THE COMMITS THE EARLIER STEPS WERE ACCEPTED AT. Their geometry is surface a
# Windows run has been produced against; P8-3 may extend the manifest but may
# not move a byte of either block.
P81_ACCEPTANCE = "35bd6ce"
P82_ACCEPTANCE = "7ff5dc7"

_CACHE: dict = {}


def _shell() -> dict:
    if "shell" not in _CACHE:
        _CACHE["shell"] = yaml.safe_load(MANIFEST.read_text(encoding="utf-8"))["phase6_shell"]
    return _CACHE["shell"]


def _projection() -> dict:
    if "charts" not in _CACHE:
        _CACHE["charts"] = json.loads(
            (BUILD / "phase8_charts_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["charts"]


def _results_projection() -> dict:
    if "results" not in _CACHE:
        _CACHE["results"] = json.loads(
            (BUILD / "phase8_results_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["results"]


def _workbook():
    if "workbook" not in _CACHE:
        manifest = json.loads((BUILD / "stage_b_manifest.json").read_text(encoding="utf-8"))
        _CACHE["workbook"] = load_workbook(BUILD / manifest["stage_a_filename"])
    return _CACHE["workbook"]


def _results():
    return _workbook()["Results"]


def _dashboard():
    return _workbook()[_projection()["chart_sheet"]]


def _charts() -> list:
    return list(_dashboard()._charts)


def _by_key() -> dict[str, dict]:
    return {chart["key"]: chart for chart in _projection()["charts"]}


def _title(chart) -> str:
    return chart.title.tx.rich.p[0].r[0].t


def _formula(address: str) -> str:
    value = _results()[address].value
    assert isinstance(value, str) and value.startswith("="), (
        f"Results!{address} holds {value!r}, not a formula")
    return value


def _block_formulas(block: str, key: str) -> list[str]:
    """Every formula in one bridge column, top to bottom."""
    entry = _projection()["bridge"][block]
    column = next(c["column"] for c in entry["columns"] if c["key"] == key)
    return [_formula(f"{column}{row}")
            for row in range(int(entry["first_row"]), int(entry["last_row"]) + 1)]


def _git(*args: str) -> str:
    import subprocess
    return subprocess.run(["git", *args], cwd=REPO_ROOT, check=True,
                          stdout=subprocess.PIPE, text=True).stdout


# ===========================================================================
# A. GENERAL - THE INVENTORY, AND WHAT IT MAY NOT DISTURB
# ===========================================================================
def test_01_the_workbook_still_has_exactly_fourteen_sheets() -> None:
    """Charts are objects anchored over cells. They need no sheet of their own,
    and a hidden chart-data sheet would be a fifteenth."""
    assert len(_workbook().sheetnames) == 14, _workbook().sheetnames


def test_02_exactly_the_four_projected_charts_exist() -> None:
    expected = _by_key()
    assert sorted(expected) == ["annual_cash_flow", "histogram", "s_curve", "tornado"]
    assert len(_charts()) == 4, len(_charts())
    titles = {_title(chart) for chart in _charts()}
    assert titles == {chart["title"] for chart in expected.values()}
    # AND NOWHERE ELSE IN THE WORKBOOK.
    for worksheet in _workbook().worksheets:
        if worksheet.title == _projection()["chart_sheet"]:
            continue
        assert not list(getattr(worksheet, "_charts", [])), worksheet.title
        assert not list(getattr(worksheet, "_images", [])), worksheet.title


def test_03_no_chart_is_three_dimensional_or_carries_junk() -> None:
    """A THIRD DIMENSION CARRIES NO DATA HERE and distorts the comparison the
    chart exists to make. A legend for one series names what the title already
    said. Category gridlines measure nothing."""
    for chart in _charts():
        assert type(chart).__name__ in ("LineChart", "BarChart"), type(chart).__name__
        assert "3D" not in type(chart).__name__
        assert chart.x_axis.majorGridlines is None
        assert chart.y_axis.majorGridlines is None
        assert chart.x_axis.delete is False and chart.y_axis.delete is False
        expected_legend = len(chart.series) > 1
        assert (chart.legend is not None) == expected_legend, (
            f"{_title(chart)} has a legend that does not match its series count")


def test_04_every_chart_sits_inside_the_reserved_region_and_below_the_summary() -> None:
    """THE P8-2 EXECUTIVE SUMMARY IS ABOVE AND STAYS ABOVE. A chart anchored one
    row too high would sit on top of a number somebody reads."""
    dashboard = _shell()["dashboard"]
    region = dashboard["chart_region"]
    summary_last = max(int(section["first_row"]) + len(section["rows"]) - 1
                       for section in dashboard["sections"])
    assert summary_last < int(region["first_row"])
    seen = set()
    for chart in _charts():
        row = chart.anchor._from.row + 1
        column = chart.anchor._from.col + 1
        assert int(region["first_row"]) <= row <= int(region["last_row"]), row
        assert row > summary_last
        assert (row, column) not in seen, "two charts share an anchor"
        seen.add((row, column))
    assert len(seen) == 4


def test_05_the_reserved_region_cells_are_still_empty() -> None:
    """A CHART IS AN OBJECT OVER CELLS, not a value in them. Anything written
    into the region would be a caption competing with the chart's own title."""
    region = _shell()["dashboard"]["chart_region"]
    sheet = _dashboard()
    for row in range(int(region["first_row"]), int(region["last_row"]) + 1):
        for column in "ABCDEFGH":
            value = sheet[f"{column}{row}"].value
            assert value is None, f"Dashboard!{column}{row} holds {value!r}"


def test_06_no_dashboard_chart_reads_machine_storage() -> None:
    """THE BRIDGE IS THE ONLY THING A CHART MAY SEE, because the bridge is where
    "" became NA(). A chart pointed at `_SimData` would plot the machine's
    blanks as zeros and would go stale the way P7-4's address did."""
    bridge_sheet = _projection()["bridge_sheet"]
    for chart in _charts():
        references = [series.val.numRef.f for series in chart.series]
        references += [series.cat.numRef.f for series in chart.series
                       if series.cat is not None and series.cat.numRef is not None]
        references += [series.cat.strRef.f for series in chart.series
                       if series.cat is not None and series.cat.strRef is not None]
        for reference in references:
            assert reference.startswith(f"'{bridge_sheet}'!"), reference
            for machine in ("_SimData", "_Calc"):
                assert machine not in reference, reference


def test_07_chart_geometry_matches_the_projection_exactly() -> None:
    """ONE AUTHORITY. Type, anchor, size, title, every series range and the
    category range - all of it answerable from the projection, and all of it
    what the workbook actually got."""
    expected = _by_key()
    by_title = {_title(chart): chart for chart in _charts()}
    kinds = {"line": ("LineChart", None), "column": ("BarChart", "col"),
             "bar": ("BarChart", "bar")}
    for spec in expected.values():
        chart = by_title[spec["title"]]
        klass, bar_type = kinds[spec["kind"]]
        assert type(chart).__name__ == klass, spec["key"]
        if bar_type is not None:
            assert chart.type == bar_type, spec["key"]
        column = re.match(r"([A-Z]+)(\d+)", spec["anchor"]).group(1)
        row = int(re.match(r"([A-Z]+)(\d+)", spec["anchor"]).group(2))
        assert chart.anchor._from.row + 1 == row
        assert chart.anchor._from.col + 1 == sum(
            (ord(ch) - 64) * 26 ** index
            for index, ch in enumerate(reversed(column)))
        extent = chart.anchor.ext
        assert round(extent.cx / 360000, 1) == round(spec["width_cm"], 1)
        assert round(extent.cy / 360000, 1) == round(spec["height_cm"], 1)
        actual_series = [series.val.numRef.f for series in chart.series]
        assert actual_series == [
            f"'{_projection()['bridge_sheet']}'!"
            + series["range"].split("!", 1)[1] for series in spec["series"]], spec["key"]
        for series in chart.series:
            reference = (series.cat.numRef.f if series.cat.numRef is not None
                         else series.cat.strRef.f)
            assert reference == (f"'{_projection()['bridge_sheet']}'!"
                                 + spec["categories"]["range"].split("!", 1)[1])


def test_08_every_bridge_block_declares_its_kind_and_its_authority() -> None:
    """DERIVED CHART-ONLY DATA OR A MIRROR - and which one it is has to be
    written down. An unlabelled helper block is an accidental second analytical
    authority waiting to be cited."""
    for name, block in _projection()["bridge"].items():
        if not isinstance(block, dict):
            continue
        assert block["kind"] in ("derived_chart_only", "mirrored"), (name, block["kind"])
        assert len(block["authority"]) > 20, name
    assert _projection()["bridge"]["status"]["kind"] == "mirrored"
    for name in ("annual", "distribution", "drivers"):
        assert _projection()["bridge"][name]["kind"] == "derived_chart_only"


def test_09_the_bridge_does_not_touch_the_accepted_p8_1_cells() -> None:
    """THOSE CELLS CARRY EVIDENCE A WINDOWS RUN WAS PRODUCED AGAINST. An overlap
    would let both sets read whatever was written last."""
    from pccm_builder.verify import _phase6_formula_cells

    spec = load_spec(MANIFEST)
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    permitted = _phase6_formula_cells(spec, structure)["Results"]
    bridge_rows = set()
    for name, block in _projection()["bridge"].items():
        if not isinstance(block, dict):
            continue
        bridge_rows |= set(range(int(block["first_row"]),
                                 int(block.get("last_row", block["first_row"])) + 1))
    accepted = _results_projection()
    accepted_rows = (
        {int(f["row"]) for f in accepted["run_stamp"]["fields"]}
        | {int(m["row"]) for m in accepted["summary"]["metrics"]}
        | set(accepted["selected"].values())
        | {int(v["row"]) for v in accepted["state"].values()}
        | set(range(int(accepted["annual"]["first_row"]),
                    int(accepted["annual"]["first_row"])
                    + int(accepted["annual"]["row_window"])))
        | set(accepted["reconciliation"]["rows"].values()))
    assert not (bridge_rows & accepted_rows), sorted(bridge_rows & accepted_rows)
    assert permitted, "the Results permitted set is empty"


def test_10_no_vba_and_no_new_state_algorithm_arrived_with_the_charts() -> None:
    """FORMULA-ONLY, and the state words still belong to their owner."""
    changed = _git("diff", "--name-only", f"{P82_ACCEPTANCE}..HEAD", "--", "pccm/src")
    assert changed.strip() == "", f"production VBA changed in P8-3: {changed}"
    for module in sorted(SRC.glob("*.bas")):
        code = "\n".join(line for line in module.read_text(encoding="utf-8").splitlines()
                         if not line.lstrip().startswith("'"))
        for word in ("Chart", "Histogram", "Tornado", "SCurve"):
            assert word not in code, f"{module.name} names {word}"


def test_11_the_accepted_p8_1_and_p8_2_geometry_did_not_move() -> None:
    """P8-3 MAY EXTEND THE MANIFEST; it may not move a byte of a surface a
    Windows run was produced against. Read from the two acceptance commits
    rather than asserted."""
    p81 = yaml.safe_load(_git("show", f"{P81_ACCEPTANCE}:pccm/spec/workbook.yaml"))
    assert p81["phase6_shell"]["results"] == _shell()["results"], (
        "the accepted P8-1 Results block changed in this step")
    p82 = yaml.safe_load(_git("show", f"{P82_ACCEPTANCE}:pccm/spec/workbook.yaml"))
    accepted = p82["phase6_shell"]["dashboard"]
    current = _shell()["dashboard"]
    for key in ("sheet", "source_sheet", "label_column", "nominal_column",
                "pv_column", "mirror_formula", "number_formats"):
        assert accepted[key] == current[key], f"the P8-2 dashboard {key} moved"
    # THE ACCEPTED SECTIONS ARE UNCHANGED; P8-3 ADDS ONE AND MOVES THE REGION.
    assert current["sections"][:len(accepted["sections"])] == accepted["sections"], (
        "an accepted P8-2 dashboard section changed")
    added = current["sections"][len(accepted["sections"]):]
    assert [section["key"] for section in added] == ["chart_status"], added


# ===========================================================================
# B. THE S-CURVE
# ===========================================================================
def _annual_guard(year: int) -> str:
    annual = _shell()["results"]["annual"]
    nominal = _shell()["results"]["nominal_column"]
    return (f'OR(${nominal}${annual["distribution_state_row"]}="NOT PRODUCED",'
            f'{year}>${nominal}${annual["year_count_row"]})')


def test_20_the_cumulative_series_is_a_running_sum_of_the_published_profile() -> None:
    """NOTHING IS RE-ALLOCATED. The S-curve is a running SUM of the annual
    profile rows the accepted P8-1 table already publishes - display arithmetic
    over cells on the same sheet, not a second stochastic allocation."""
    results = _shell()["results"]["annual"]
    first = int(results["first_row"])
    display = {c["field"]: c["column"] for c in results["columns"]}
    for measure, column in (("nominal", display["nominal"]), ("pv", display["pv"])):
        formulas = _block_formulas("annual", f"cumulative_{measure}")
        for offset, formula in enumerate(formulas):
            span = f"SUM(${column}${first}:${column}${first + offset})"
            assert span in formula, (offset, measure, formula)
            assert _annual_guard(offset + 1) in formula, (offset, formula)
    # AND THE TWO MEASURES ARE NOT SWAPPED.
    nominal = _block_formulas("annual", "cumulative_nominal")[0]
    pv = _block_formulas("annual", "cumulative_pv")[0]
    assert f"${display['nominal']}$" in nominal and f"${display['pv']}$" not in nominal
    assert f"${display['pv']}$" in pv and f"${display['nominal']}$" not in pv


def test_21_the_annual_series_is_bounded_by_the_stamped_year_count() -> None:
    """NO POST-WINDOW ZERO YEARS. Every row past the stamped count returns NA(),
    which every chart type declines to plot. `""` would be drawn as zero - a
    fabricated year of a project costing nothing - which is the exact reason
    this bridge exists rather than the chart reading the table."""
    for key in ("project_index", "calendar_year", "annual_nominal",
                "cumulative_nominal", "cumulative_pv"):
        formulas = _block_formulas("annual", key)
        assert len(formulas) == _projection()["bridge"]["annual"]["row_count"]
        for offset, formula in enumerate(formulas):
            assert formula.startswith("=IF("), formula
            assert ",NA()," in formula, (key, offset, formula)
            assert _annual_guard(offset + 1) in formula, (key, offset)
            assert '""' not in formula, (
                f"{key} row {offset} can return a blank a chart would plot as zero")


def test_22_the_s_curve_plots_the_cumulative_series_against_calendar_year() -> None:
    spec = _by_key()["s_curve"]
    assert spec["kind"] == "line"
    assert [series["key"] for series in spec["series"]] == [
        "cumulative_nominal", "cumulative_pv"]
    assert spec["categories"]["key"] == "calendar_year"
    assert spec["source_block"] == "annual"
    assert spec["state_source"] == "profile_state"
    # A LINE CHART OF A CUMULATIVE SERIES IS NOT SMOOTHED. Interpolating between
    # two published years would draw money that was never allocated.
    for chart in _charts():
        if _title(chart) != spec["title"]:
            continue
        for series in chart.series:
            assert series.smooth is False or series.smooth is None


# ===========================================================================
# C. THE HISTOGRAM
# ===========================================================================
def test_30_the_bin_contract_is_deterministic_and_projected() -> None:
    """A FIXED COUNT, NOT A RULE THAT READS THE DATA. The same run always bins
    the same way, so two readers never see two histograms of one result."""
    contract = _projection()["bin_contract"]
    assert contract["bin_count"] == _projection()["bridge"]["distribution"]["row_count"]
    assert contract["bin_count"] >= 2
    assert contract["measure"] == "nominal"
    for clause in ("equal width", "[lower, upper)", "<= upper", "NA()"):
        assert clause in json.dumps(contract), clause


def test_31_the_edges_come_from_the_published_minimum_and_maximum() -> None:
    """NO STATISTIC IS RECOMPUTED. The extremes are the ones the run published
    on Results; a bridge that scanned the iteration column for its own min and
    max would be a second implementation of two numbers already settled."""
    summary = {m["key"]: m["row"] for m in _results_projection()["summary"]["metrics"]}
    nominal = _results_projection()["columns"]["nominal"]
    minimum = f"${nominal}${summary['minimum']}"
    maximum = f"${nominal}${summary['maximum']}"
    lowers = _block_formulas("distribution", "lower")
    uppers = _block_formulas("distribution", "upper")
    count = _projection()["bin_contract"]["bin_count"]
    for index, (lower, upper) in enumerate(zip(lowers, uppers)):
        width = f"(({maximum}-{minimum})/{count})"
        assert f"({minimum}+{index}*{width})" in lower, (index, lower)
        assert f"({minimum}+{index + 1}*{width})" in upper, (index, upper)
    assert "MIN(" not in "".join(lowers) and "MAX(" not in "".join(uppers)
    assert "PERCENTILE" not in "".join(lowers + uppers)


def test_32_every_iteration_lands_in_exactly_one_bin() -> None:
    """COUNT CONSERVATION, READ OFF THE COMPARATORS. Bins are half-open so no
    iteration is counted twice, and the LAST one closes so the maximum is
    counted at all rather than falling off the end."""
    counts = _block_formulas("distribution", "count")
    for index, formula in enumerate(counts[:-1]):
        assert '"<"&' in formula, (index, "not half-open")
        assert '"<="&' not in formula, (index, "double-counts its upper edge")
        assert '">="&' in formula
    assert '"<="&' in counts[-1], "the last bin drops the maximum"
    assert '">="&' in counts[-1]


def test_33_an_unpublished_or_degenerate_run_fabricates_no_bin() -> None:
    """A HISTOGRAM OF ZEROS IS A CONFIDENT PICTURE OF A RUN THAT NEVER HAPPENED.
    And a run whose iterations all produced the same total has no width to bin:
    the whole population goes in the first bin and the rest return NA()."""
    stamp = {f["key"]: f["row"] for f in _results_projection()["run_stamp"]["fields"]}
    nominal = _results_projection()["columns"]["nominal"]
    published = f"${nominal}${stamp['run_id']}"
    counts = _block_formulas("distribution", "count")
    for index, formula in enumerate(counts):
        assert formula.startswith(f'=IF({published}="",NA(),'), (index, formula)
        if index == 0:
            assert f"${nominal}${stamp['iterations_run']}" in formula
        else:
            assert ",NA())" in formula or ",NA()," in formula, index
    for formula in _block_formulas("distribution", "lower"):
        assert formula.startswith(f'=IF({published}="",NA(),')


def test_34_the_histogram_plots_the_count_against_the_lower_edge() -> None:
    spec = _by_key()["histogram"]
    assert spec["kind"] == "column"
    assert [series["key"] for series in spec["series"]] == ["count"]
    assert spec["categories"]["key"] == "lower"
    assert spec["state_source"] == "distribution_state"
    assert spec["source_block"] == "distribution"


# ===========================================================================
# D. THE ANNUAL CASH FLOW
# ===========================================================================
def test_40_the_annual_chart_tracks_the_published_profile_exactly() -> None:
    """THE PERSISTED ANSWER, NOT A NEW ONE. Each bar is the profile cell the
    accepted P8-1 table publishes for that year, guarded and nothing else."""
    display = {c["field"]: c["column"] for c in _shell()["results"]["annual"]["columns"]}
    first = int(_shell()["results"]["annual"]["first_row"])
    for offset, formula in enumerate(_block_formulas("annual", "annual_nominal")):
        assert f"${display['nominal']}${first + offset}" in formula, (offset, formula)
        assert "SUM(" not in formula, "the annual bar is a running total"
    spec = _by_key()["annual_cash_flow"]
    assert spec["kind"] == "column"
    assert [series["key"] for series in spec["series"]] == ["annual_nominal"]
    assert spec["categories"]["key"] == "calendar_year"
    assert spec["state_source"] == "profile_state"


def test_41_the_year_ordering_is_the_published_one() -> None:
    """ROW ORDER IS YEAR ORDER, and it is taken rather than reconstructed: the
    bridge row for year n reads the display row for year n."""
    first = int(_shell()["results"]["annual"]["first_row"])
    display = {c["field"]: c["column"] for c in _shell()["results"]["annual"]["columns"]}
    for offset, formula in enumerate(_block_formulas("annual", "calendar_year")):
        assert f"${display['calendar_year']}${first + offset}" in formula, offset
    for offset, formula in enumerate(_block_formulas("annual", "project_index")):
        assert f"${display['project_index']}${first + offset}" in formula, offset


# ===========================================================================
# E. THE TORNADO
# ===========================================================================
def test_50_the_tornado_selects_a_window_and_ranks_nothing() -> None:
    """PHASE 7 OWNS THE RANKING. The Sensitivity sheet publishes it in rank
    order, so the first N rows ARE the top N - taking them preserves the order
    by construction, and there is nothing here that could re-sort it."""
    sensitivity = _shell()["sensitivity"]
    columns = {c["key"]: c["column"] for c in sensitivity["columns"]}
    first = int(sensitivity["first_row"])
    names = _block_formulas("drivers", "driver_name")
    rhos = _block_formulas("drivers", "rho")
    assert len(names) == len(rhos) == _projection()["bridge"]["drivers"]["row_count"]
    for index, (name, rho) in enumerate(zip(names, rhos)):
        assert f"{sensitivity['sheet']}!${columns['driver_name']}${first + index}" in name
        assert f"{sensitivity['sheet']}!${columns['rho']}${first + index}" in rho
    # NOT ONE SORTING, RANKING OR ABSOLUTE-VALUE CONSTRUCT ANYWHERE.
    for formula in names + rhos:
        for banned in ("SORT", "RANK", "LARGE", "SMALL", "ABS(", "INDEX", "MATCH"):
            assert banned not in formula, f"the tornado re-derives its order: {formula}"


def test_51_the_signed_rho_is_preserved_and_the_abs_column_is_not_plotted() -> None:
    """DIRECTION IS THE POINT OF A TORNADO. Plotting |rho| would draw every bar
    to the right and lose which drivers push the total up and which pull it
    down - and the Sensitivity sheet publishes both columns, so the wrong one is
    one character away."""
    sensitivity = _shell()["sensitivity"]
    columns = {c["key"]: c["column"] for c in sensitivity["columns"]}
    assert "abs_rho" in columns, "the accepted sheet no longer publishes |rho|"
    for formula in _block_formulas("drivers", "rho"):
        assert f"${columns['abs_rho']}$" not in formula, (
            "the tornado plots the absolute value and loses direction")
        assert f"${columns['rho']}$" in formula
    spec = _by_key()["tornado"]
    assert [series["key"] for series in spec["series"]] == ["rho"]
    assert spec["kind"] == "bar", "a tornado is horizontal"
    assert spec["categories"]["key"] == "driver_name"


def test_52_absent_drivers_are_not_drawn_as_zero_length_bars() -> None:
    """FEWER ELIGIBLE DRIVERS THAN N MUST DRAW FEWER BARS. N-k bars of nothing
    sitting on the axis look measured, and zero-variance drivers are excluded by
    the accepted contract precisely so they are not shown as uncorrelated."""
    for formula in _block_formulas("drivers", "driver_name") + _block_formulas("drivers", "rho"):
        assert '="",NA()' in formula.replace(" ", ""), formula
        assert ',""' not in formula, "an absent driver becomes a blank a chart plots as zero"


def test_53_an_unavailable_ranking_is_stated_rather_than_drawn_empty() -> None:
    """A CONVINCING EMPTY TORNADO IS THE WORST OUTCOME. The Sensitivity sheet
    already writes a sentence saying whether its ranking exists and which run it
    belongs to; the bridge mirrors that sentence and the Dashboard shows it."""
    status = _projection()["bridge"]["status"]
    assert status["kind"] == "mirrored"
    keys = {row["key"] for row in status["rows"]}
    assert "sensitivity_availability" in keys
    sensitivity = _shell()["sensitivity"]
    availability = (f"{sensitivity['sheet']}!"
                    f"${sensitivity['columns'][1]['column']}$"
                    f"{sensitivity['availability_row']}")
    formula = _formula(f"{_shell()['results']['nominal_column']}{status['first_row']}")
    assert availability in formula, formula
    # MIRRORED, NEVER RE-DERIVED: no arm of that sentence is rebuilt here.
    for banned in ("PUBLISHED", "CURRENT", "IF(AND(", "COUNTIF"):
        assert banned not in formula, f"the bridge re-derives the availability: {banned}"
    # AND THE DASHBOARD SHOWS IT.
    dashboard = json.loads(
        (BUILD / "phase8_dashboard_inspection.json").read_text(encoding="utf-8"))
    mirrored = {(entry["source_block"], entry["source_key"])
                for section in dashboard["sections"] for entry in section["rows"]}
    assert ("chart_status", "sensitivity_availability") in mirrored
    assert _by_key()["tornado"]["state_source"] == "sensitivity_availability"


# ===========================================================================
# F. STATE - NO CHART OWNS ONE
# ===========================================================================
def test_60_every_chart_names_a_state_source_it_does_not_own() -> None:
    """A CHART CARRIES NO STATE LOGIC. It inherits the state its source already
    publishes, and the projection says where a reader finds it."""
    published = {"distribution_state", "profile_state", "profile_px",
                 "sensitivity_availability"}
    for spec in _by_key().values():
        assert spec["state_source"] in published, spec
        assert spec["no_data_value"] == "NA()"


def test_61_no_chart_bridge_formula_carries_a_state_word_of_its_own() -> None:
    """THE ONE STATE WORD THE BRIDGE MAY SPELL is the NOT PRODUCED the accepted
    annual table already guards on - and it spells it because it reuses that
    exact guard. Every other word would be a second vocabulary."""
    everything = []
    for block in ("annual", "distribution", "drivers"):
        entry = _projection()["bridge"][block]
        for column in entry["columns"]:
            everything += _block_formulas(block, column["key"])
    for formula in everything:
        for word in ("CURRENT", "HISTORICAL", "OTHER Px", "Reconciled",
                     "NOT RECONCILED", "STALE", "INVALID"):
            assert word not in formula, f"a bridge formula spells {word!r}: {formula}"
    assert any("NOT PRODUCED" in formula for formula in everything), (
        "the annual bridge no longer reuses the accepted empty-state guard")


# ===========================================================================
# G. MUTATIONS - each is a way a chart lies convincingly
# ===========================================================================
def _mutated(mutate) -> Path:
    raw = yaml.safe_load(MANIFEST.read_text(encoding="utf-8"))
    mutate(raw["phase6_shell"]["charts"], raw["phase6_shell"])
    path = BUILD / "_phase8_charts_probe.yaml"
    path.write_text(yaml.safe_dump(raw, sort_keys=False), encoding="utf-8")
    return path


@pytest.mark.parametrize("name,mutate,expected", [
    # THE WRONG SOURCE SHEET - the bridge no longer lives where the charts read.
    ("the bridge moves off Results",
     lambda charts, shell: charts.__setitem__("bridge_sheet", "Sensitivity"),
     "the bridge lives on"),
    # A SERIES NOTHING PUBLISHES - the P7-4 stale-address shape.
    ("a chart plots a series the bridge does not publish",
     lambda charts, shell: charts["charts"][0]["series"].__setitem__(
         0, {"key": "expected_exposure", "name": "Expected"}),
     "which the annual block does not publish"),
    # THE CATEGORY TAKEN FROM A BLOCK THAT HAS NO SUCH COLUMN.
    ("a chart takes categories from a column that does not exist",
     lambda charts, shell: charts["charts"][1].__setitem__("categories", "calendar_year"),
     "which the distribution block does not publish"),
    # A THIRD DIMENSION.
    ("a chart becomes three-dimensional",
     lambda charts, shell: charts["charts"][0].__setitem__("kind", "bar3D"),
     "is a 'bar3D' chart"),
    # THE BRIDGE GROWN UP INTO THE ACCEPTED P8-1 SURFACE.
    ("the bridge starts inside the accepted Results layout",
     lambda charts, shell: charts["bridge"].__setitem__("heading_row", 100),
     "inside the accepted Results layout"),
    # TWO BRIDGE BLOCKS OVERLAID - the silent one.
    ("two bridge blocks write the same row",
     lambda charts, shell: charts["bridge"]["drivers"].__setitem__(
         "first_row", charts["bridge"]["distribution"]["first_row"]),
     "both write row"),
    # A CHART ANCHORED OVER THE EXECUTIVE SUMMARY.
    ("a chart is anchored above the reserved region",
     lambda charts, shell: charts["charts"][0].__setitem__("anchor", "B20"),
     "outside the reserved region"),
    # A STATE SOURCE NOTHING PUBLISHES.
    ("a chart names a state source nothing publishes",
     lambda charts, shell: charts["charts"][0].__setitem__("state_source", "model_status"),
     "which nothing publishes"),
])
def test_70_the_manifest_refuses_each_way_a_chart_lies(
        name: str, mutate, expected: str) -> None:
    """EACH MUTATION APPLIED TO A COPY OF THE MANIFEST and refused by the loader,
    naming what is wrong. Nothing on disk changes."""
    path = _mutated(mutate)
    try:
        with pytest.raises(SpecError, match=re.escape(expected)):
            load_spec(path)
    finally:
        path.unlink(missing_ok=True)


@pytest.mark.parametrize("name,mutate", [
    # THE TWO LADDERS, ONE SHEET OUT: a chart plotting its own axis.
    ("a chart plots its category column as a series",
     lambda inspection: inspection["charts"][0]["series"].__setitem__(
         0, dict(inspection["charts"][0]["categories"], name="Category"))),
    # A RANGE THAT LOST ITS SHEET - it would resolve against the Dashboard,
    # which holds none of this data, and plot nothing while looking fine.
    ("a series range loses its sheet",
     lambda inspection: inspection["charts"][0]["series"][0].__setitem__(
         "range", "$H$279:$H$478")),
    # A LEGEND THAT DISAGREES WITH THE SERIES COUNT.
    ("a single-series chart declares a legend",
     lambda inspection: inspection["charts"][1].__setitem__("legend", True)),
    # TWO CHARTS IN ONE PLACE.
    ("two charts share an anchor",
     lambda inspection: inspection["charts"][1].__setitem__(
         "anchor", inspection["charts"][0]["anchor"])),
])
def test_71_the_projection_validator_refuses_each_incoherent_chart(
        name: str, mutate) -> None:
    """THE LOADER CANNOT SEE THESE - they are structurally legal manifests whose
    PROJECTION is incoherent - so the projection's own validator must."""
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    inspection = build_phase8_charts_inspection(
        load_spec(MANIFEST), structure.limits.max_generated_year_columns)
    mutate(inspection)
    with pytest.raises(ValueError):
        validate_phase8_charts_inspection(inspection)


def test_72_the_unmutated_manifest_passes_both_gates() -> None:
    """SO THE TWELVE REFUSALS ABOVE ARE REFUSALS OF THE MUTATION, not of the
    fixture."""
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    inspection = build_phase8_charts_inspection(
        load_spec(MANIFEST), structure.limits.max_generated_year_columns)
    validate_phase8_charts_inspection(inspection)
    assert inspection == _projection(), (
        "the committed chart projection is not what the manifest now produces")
