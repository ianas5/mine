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
# THE ACCEPTED CHART PACKAGE, before the value-axis bounds were taken off.
ACCEPTED_CHART_PACKAGE = "e87d566"
# The PowerShell the executed harnesses run under, as the other suites name it.
PWSH = "/opt/pwsh/pwsh"

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


def _chart_module() -> str:
    """The Dashboard chart presentation owner's source."""
    return (PCCM_ROOT / "src" / "vba" / "modChartPresentation.bas").read_text(encoding="utf-8")


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
def _accepted_series_parts(title: str) -> list[tuple[str, str]]:
    """The same, from a workbook built by the ACCEPTED pre-polish builder at
    ebeae65 - the tree whose literal-range charts Windows proved at P8-3."""
    if "accepted_parts" not in _CACHE:
        import subprocess as _sp
        import tempfile
        root = Path(tempfile.mkdtemp(prefix="pccm-accepted-"))
        _sp.run(["git", "worktree", "add", "-f", "--detach", str(root), PREVIOUS_HEAD],
                cwd=PCCM_ROOT.parent, check=True, stdout=_sp.DEVNULL, stderr=_sp.DEVNULL)
        try:
            _sp.run([sys.executable, str(root / "pccm" / "builder" / "build_stage_a.py"), "--quiet"],
                    check=True, stdout=_sp.DEVNULL)
            built = root / "pccm" / "build" / "PCCM_stageA.xlsx"
            parts = {}
            with zipfile.ZipFile(built) as archive:
                for name in archive.namelist():
                    if re.fullmatch(r"xl/charts/chart\d+\.xml", name):
                        text = archive.read(name).decode("utf-8")
                        found = re.search(r"<a:t>([^<]+)</a:t>", text)
                        if found:
                            parts[found.group(1)] = text
            _CACHE["accepted_parts"] = parts
        finally:
            _sp.run(["git", "worktree", "remove", "--force", str(root)],
                    cwd=PCCM_ROOT.parent, stdout=_sp.DEVNULL, stderr=_sp.DEVNULL)
    part = _CACHE["accepted_parts"][title]
    out = []
    for block in re.findall(r"<ser>.*?</ser>", part, re.S):
        category = re.search(r"<cat>(.*?)</cat>", block, re.S)
        value = re.search(r"<val>(.*?)</val>", block, re.S)
        assert category is not None and value is not None, title
        out.append((category.group(1), value.group(1)))
    assert out, title
    return out


def _series_parts(title: str) -> list[tuple[str, str]]:
    """Every series of one chart, as (category reference, value reference) read
    out of the chart part itself - the bytes Excel is handed."""
    part = _part_for(title)
    out = []
    for block in re.findall(r"<ser>.*?</ser>", part, re.S):
        category = re.search(r"<cat>(.*?)</cat>", block, re.S)
        value = re.search(r"<val>(.*?)</val>", block, re.S)
        assert category is not None, f"{title}: a series carries no category source"
        assert value is not None, f"{title}: a series carries no value source"
        out.append((category.group(1), value.group(1)))
    assert out, title
    return out


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


def test_02_the_year_charts_plot_the_applied_years_through_the_bound_value_names() -> None:
    """THE VALUE HALF OF THE CORRECTION, WHICH WINDOWS HAS TWICE PROVED. Every
    value series of both year charts reads its applied-year name, never the
    200-row window - in the projection and in the chart parts the workbook
    actually carries. The category half is test_02b's: it is bound at runtime,
    because a named category comes back blank."""
    sheet = _projection()["bridge_sheet"]
    block = _projection()["bridge"]["annual"]
    binding = block["applied_binding"]
    assert binding["name_prefix"] == "chartAnnual" and binding["extent"] == "year_count"
    assert binding["window_rows"] == _window() == block["row_count"]
    names = {column["key"]: column["applied_binding"] for column in block["columns"]
             if column["applied_binding"]}
    for key in YEAR_CHARTS:
        spec = _by_key()[key]
        chart = _chart_objects()[spec["title"]]
        for entry in spec["series"]:
            bound = names[entry["key"]]
            assert bound["name"] == f"chartAnnual_{entry['key']}"
            assert entry["binding"] == bound["name"]
            assert entry["range"] == f"{sheet}!{bound['name']}"
            assert entry["window_range"] == bound["window_range"]
            assert entry["range"] != entry["window_range"]
            # A SERIES BOUND TO THE RESERVED TAIL WOULD END AT ROW 478.
            assert re.search(r"\$\d+:\$[A-Z]+\$\d+$", entry["range"]) is None
        assert [s.val.numRef.f for s in chart.series] == [
            f"'{sheet}'!chartAnnual_{s['key']}" for s in spec["series"]]

def test_02b_every_category_is_a_literal_range_and_the_year_charts_are_built_on_one_row() -> None:
    """THREE WINDOWS RUNS SETTLED THIS. A defined name in a series' CATEGORY
    slot comes back blank whatever the markup around it - first with the two
    year charts alone, then, after a cache was added to every named reference,
    with all four - and the axis is numbered 1, 2, 3. Excel binds a category
    from a RANGE, at runtime, through Series.XValues, which Windows proved
    directly by assigning one and reading the SERIES formula back.

    So no chart is BUILT on a name any more. The two year charts are built on
    the FIRST reserved annual row - a literal range Excel resolves while it
    reads the chart part, and the right picture for a workbook with nothing
    published - and modChartPresentation widens that binding to the published
    year count. The histogram and the tornado keep their whole block, and their
    markup is required byte-identical to a build from the ebeae65 worktree.

    THE VALUE NAMES STAY. Windows has twice proved those survive."""
    runtime = _projection()["bridge"]["annual"]["runtime_category"]
    for key in YEAR_CHARTS:
        spec = _by_key()[key]
        for category, value in _series_parts(spec["title"]):
            # THE CATEGORY: a literal cell, no name, no cache.
            formula = re.search(r"<f>(.*?)</f>", category, re.S).group(1)
            assert formula == "'Results'!$D$279", (key, formula)
            assert "chartAnnual" not in category, (key, "the failed named category is back")
            assert "<numCache>" not in category, (key, "the withdrawn cache is back")
            # THE VALUE: the applied-year name, unchanged and uncached.
            value_formula = re.search(r"<f>(.*?)</f>", value, re.S).group(1)
            assert value_formula.startswith("'Results'!chartAnnual_"), (key, value_formula)
            assert "<numCache>" not in value, (key, "the withdrawn cache is back")
            assert re.search(r"\$[A-Z]+\$\d+", value_formula) is None, (key, value_formula)
        assert spec["categories"]["range"] == runtime["built_range"]
        assert spec["categories"]["binding"] is None
        assert spec["categories"]["runtime_binding"] == runtime["window_name"]
    # NO 200-ROW VISIBLE CATEGORY BINDING ANYWHERE IN THE FILE.
    for part in _chart_parts():
        assert "$D$478" not in part, "a chart is built on the whole reserved window"
    # THE CATEGORY IS NEVER ABSENT AND NEVER EMPTY - the shape Excel left behind.
    for key in YEAR_CHARTS:
        part = _part_for(_by_key()[key]["title"])
        assert "<cat />" not in part and "<cat/>" not in part, key
        assert part.count("<ser>") == part.count("<cat>") == len(_by_key()[key]["series"]), key
    # THE TWO RANGE-BOUND CHARTS: byte-identical to the accepted build, series
    # for series - the markup P8-3 proved, untouched by any of this.
    for key in ("histogram", "tornado"):
        title = _by_key()[key]["title"]
        assert _series_parts(title) == _accepted_series_parts(title), (
            f"{title}: the literal category path is not the accepted one")
        for category, value in _series_parts(title):
            for source in (category, value):
                assert "<numCache>" not in source, (key, "a plain range gained a cache")
                assert re.search(r"<f>'[A-Za-z_]+'!\$[A-Z]+\$\d+:\$[A-Z]+\$\d+</f>", source), (key, source)
        assert _by_key()[key]["categories"]["runtime_binding"] is None, key
    # AND EVERY CATEGORY GOES THROUGH THE ACCEPTED CALL PATH.
    builder = (PCCM_ROOT / "builder" / "pccm_builder" / "workbook_builder.py").read_text(encoding="utf-8")
    assert builder.count("chart.set_categories(categories)") == 1
    assert "def _chart_series(values: str, title: str):" in builder
    assert "item.cat = " not in builder, "a category is assembled by hand again"

def test_03_the_names_are_workbook_scoped_and_read_the_one_year_count_cell() -> None:
    """ONE YEAR-COUNT AUTHORITY, AND WORKBOOK SCOPE. Every applied-year VALUE
    name cuts its column at the Results annual state cell the bridge guard
    already reads, and the two names the presentation owner reads point at the
    same column window and the same cell. The names are defined at WORKBOOK
    level and referenced through the sheet, `Results!<name>`, never through a
    file name, so nothing in a reference changes when Stage A is saved as
    Stage B or the distribution copy is saved under another name.

    AND THE FAILED CATEGORY NAME IS GONE. The calendar-year column carries no
    cut name at all now; it is bound at runtime from a range."""
    results = _workbook()[_projection()["bridge_sheet"]]
    block = _projection()["bridge"]["annual"]
    binding = block["applied_binding"]
    runtime = block["runtime_category"]
    shell = _shell()
    expected_extent = "Results!" + _annual_extent_cell(shell["results"], "year_count")
    assert binding["extent_cell"] == expected_extent
    assert runtime["extent_formula"] == expected_extent, "a second year-count authority"
    assert binding["runtime_category"] == runtime["key"] == "calendar_year"
    # THE GUARD READS THE SAME CELL: `1>$D$56` at the first bridge row.
    first = int(block["first_row"])
    guard_cell = expected_extent.split("!", 1)[1]
    for column in block["columns"]:
        formula = results[f"{column['column']}{first}"].value
        assert f",1>{guard_cell})" in formula, (column["key"], formula)
    # THE NAMES, AT WORKBOOK LEVEL, EXACTLY AS PROJECTED - and none of them left
    # behind on the sheet, where a sheet-qualified reference would find the
    # local one first.
    named = {column["key"]: column["applied_binding"] for column in block["columns"]
             if column["applied_binding"]}
    assert runtime["key"] not in named, "the failed category name is still written"
    in_book = {name: entry.attr_text for name, entry in _workbook().defined_names.items()}
    assert not any(name.startswith("chartAnnual") for name in results.defined_names)
    for key, bound in named.items():
        assert in_book[bound["name"]] == bound["formula"]
        assert bound["scope"] == "workbook"
        assert bound["reference"] == f"{_projection()['bridge_sheet']}!{bound['name']}"
        assert "." not in bound["reference"] and ".xls" not in bound["reference"]
        assert _NAME_FORMULA.fullmatch(bound["formula"]), bound["formula"]
    # THE TWO THE PRESENTATION OWNER READS: the reserved column window, and the
    # Years Covered cell - neither cut, neither computed.
    assert in_book[runtime["window_name"]] == runtime["window_formula"] == "Results!$D$279:$D$478"
    assert in_book[runtime["extent_name"]] == runtime["extent_formula"]
    assert runtime["window_formula"] == next(
        c["range"] for c in block["columns"] if c["key"] == runtime["key"])
    assert "INDEX" not in in_book[runtime["window_name"]]
    assert "INDEX" not in in_book[runtime["extent_name"]]
    # AND THE BUILDER'S OWN STRUCTURAL VERIFICATION DERIVES THEM FROM THE
    # MANIFEST rather than listing them.
    verify = (PCCM_ROOT / "builder" / "pccm_builder" / "verify.py").read_text(encoding="utf-8")
    assert "chart_names = set(_applied_year_chart_names(spec, structure))" in verify
    assert "not (found_names - expected_names - chart_names)" in verify
    assert "and not (expected_names - found_names)" in verify
    assert "chartAnnual" not in verify, "the verification lists a name it should derive"

# THE PRESENTATION OWNER'S OWN ARITHMETIC, in Python: what
# modChartPresentation.CategoryRange resolves for one reported year count.
def _runtime_rows(year_count, window: int) -> int:
    if isinstance(year_count, bool) or not isinstance(year_count, (int, float)):
        return 1
    if float(year_count) < 1:
        return 1
    if float(year_count) >= float(window):
        return window
    return int(year_count)


@pytest.mark.parametrize("year_count,expected_rows", [
    (1, 1), (3, 3), (7, 7), (10, 10), (30, 30), (75, 75), (200, 200),
    # OUTSIDE THE WINDOW OR NOT A COUNT: the binding stays inside the reserved
    # column and never below one row, whose label the bridge answers blank and
    # whose value it answers NA(). No year is invented and no point is created.
    (250, 200), (0, 1), ("", 1), ("NOT PRODUCED", 1),
])
def test_04_every_applied_duration_binds_exactly_its_own_category_rows(year_count, expected_rows) -> None:
    """MULTIPLE DURATIONS, NOT THE DEMO'S. For each stamped year count the
    presentation owner resizes the reserved calendar-year column to exactly
    that many rows, starting at the block's first row and never past its last,
    and the bridge guard turns every value beyond it to NA() and every label to
    blank - so the binding and the data agree row for row.

    THE ARITHMETIC IS READ OFF THE MODULE, not assumed: the bounds below are
    the ones its own source states."""
    block = _projection()["bridge"]["annual"]
    runtime = block["runtime_category"]
    window = int(block["row_count"])
    first = int(block["first_row"])
    rows = _runtime_rows(year_count, window)
    assert rows == expected_rows, (year_count, rows)
    assert first == int(runtime["first_row"])
    assert first + rows - 1 <= int(runtime["last_row"])
    # THE MODULE'S OWN BOUNDS, in its own words.
    module = _chart_module()
    assert "If CDbl(reported) >= 1 Then" in module
    assert "If CDbl(reported) >= CDbl(reserved.Rows.Count) Then" in module
    assert "years = reserved.Rows.Count" in module
    assert "years = CLng(Int(CDbl(reported)))" in module
    assert "years = 1" in module and "If IsNumeric(reported) Then" in module
    assert "Set categories = reserved.Resize(years, 1)" in module
    # AND THE GUARD BEYOND THE SELECTION SAYS THE SAME THING.
    results = _workbook()[_projection()["bridge_sheet"]]
    guard_cell = runtime["extent_formula"].split("!", 1)[1]
    if isinstance(year_count, int) and not isinstance(year_count, bool) and 0 < year_count < window:
        for column in block["columns"]:
            beyond = results[f"{column['column']}{first + rows}"].value
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
        # A VALUE SERIES BOUND BACK TO THE WHOLE WINDOW.
        damaged = copy.deepcopy(inspection)
        chart = next(c for c in damaged["charts"] if c["key"] == key)
        chart["series"][0]["range"] = chart["series"][0]["window_range"]
        chart["series"][0]["binding"] = None
        with pytest.raises(ValueError, match="applied-year name|whole reserved window|whole window"):
            validate_phase8_charts_inspection(damaged)
        # THE CATEGORY BUILT ON THE WHOLE WINDOW instead of the first row - the
        # 200 crowded slots this correction exists to remove.
        damaged = copy.deepcopy(inspection)
        chart = next(c for c in damaged["charts"] if c["key"] == key)
        chart["categories"]["range"] = chart["categories"]["window_range"]
        with pytest.raises(ValueError, match="not the first reserved row"):
            validate_phase8_charts_inspection(damaged)
        # OR BOUND THROUGH A NAME AGAIN, which is what Excel returns blank.
        damaged = copy.deepcopy(inspection)
        chart = next(c for c in damaged["charts"] if c["key"] == key)
        chart["categories"]["binding"] = "chartAnnual_calendar_year"
        with pytest.raises(ValueError, match="Excel returns a named category blank"):
            validate_phase8_charts_inspection(damaged)
        # OR CUT AT A CELL THAT IS NOT THE ONE YEAR-COUNT AUTHORITY.
        damaged = copy.deepcopy(inspection)
        damaged["bridge"]["annual"]["runtime_category"]["extent_formula"] = "Results!$D$57"
        with pytest.raises(ValueError, match="not the block's extent cell"):
            validate_phase8_charts_inspection(damaged)
    # A NAME THAT DOES NOT CUT AT THE EXTENT CELL, OR OUTSIDE THE WINDOW.
    for mutate in (
        lambda b: b.__setitem__("formula", b["formula"].replace("MIN(200,", "MIN(250,")),
        lambda b: b.__setitem__("formula", b["formula"].replace("MAX(1,", "MAX(0,")),
        lambda b: b.__setitem__("formula", b["formula"].replace("$D$56", "$D$57")),
        lambda b: b.__setitem__("extent_cell", "Results!$D$57"),
    ):
        damaged = copy.deepcopy(inspection)
        plotted = {series["key"] for chart in damaged["charts"]
                   if chart["source_block"] == "annual" for series in chart["series"]}
        named = next(column for column in damaged["bridge"]["annual"]["columns"]
                     if column["applied_binding"] and column["key"] in plotted)
        mutate(named["applied_binding"])
        with pytest.raises(ValueError):
            validate_phase8_charts_inspection(damaged)
    # AND THE MANIFEST WITHOUT THE BINDING BUILDS WINDOW-BOUND CHARTS AGAIN,
    # which is exactly the shape the projection's own validator accepts as
    # "no binding declared" - so the binding's presence is pinned here.
    assert _shell()["charts"]["bridge"]["annual"]["applied_binding"] == {
        "name_prefix": "chartAnnual", "extent": "year_count",
        "runtime_category": "calendar_year"}


def test_06_the_manifest_refuses_a_binding_it_cannot_resolve() -> None:
    text = MANIFEST.read_text(encoding="utf-8")
    for old, new, message in (
        ('          extent: "year_count"\n', '          extent: "years_stamped"\n',
         "not a row the Results annual block publishes"),
        ('          name_prefix: "chartAnnual"\n', '          name_prefix: "9charts"\n',
         "cannot begin a defined name"),
        ("          major_unit: 0.10\n", "          major_unit: 0\n", "is not a step"),
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
INSPECTOR = PCCM_ROOT / "bootstrap" / "windows" / "phase10_chart_polish_inspect.ps1"
GATE_HARNESS = PCCM_ROOT / "tests" / "phase10_chart_polish_gate_flow.ps1"


def _gate_lines() -> dict[str, str]:
    done = subprocess.run([PWSH, "-NoProfile", "-File", str(GATE_HARNESS),
                           "-Inspector", str(INSPECTOR)],
                          capture_output=True, text=True, timeout=300)
    assert done.returncode == 0, done.stdout + done.stderr
    lines = {}
    for line in done.stdout.splitlines():
        parts = line.split("|", 2)
        if len(parts) == 3 and parts[0] == "GATE":
            lines[parts[1]] = parts[2]
    assert lines, done.stdout
    return lines


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_05b_the_windows_gate_judges_the_live_xvalues_payload() -> None:
    """EXECUTED, AND THE ORACLE HAS MOVED. Windows proved the runtime binding
    works - a disposable copy returned XVALUES|2099 with the Dashboard
    protected again - and proved at the same time that the SERIES formula's
    category argument can read blank for a series whose categories are plainly
    present, which is why this gate once reported the histogram and the tornado
    broken while their accepted populations were intact. So the verdict is
    taken on Series.XValues, normalised from whatever COM returns and compared
    against the cells the contract expects, in order.

    The formula is still printed, as a diagnostic, and is never the oracle."""
    lines = _gate_lines()
    # 1 and 3. NOTHING PUBLISHED: exactly one category, the first bridge cell's
    #    own value, which is blank - from a scalar and from a one-item array.
    for case in ("year.nothing-published.scalar", "year.nothing-published.array"):
        assert lines[case].startswith("count=1|payload=<blank>|"), (case, lines[case])
        assert "|verdict=accepted|" in lines[case], (case, lines[case])
    # 3 and 4. ONE YEAR, N YEARS, and the whole window when the whole window is
    #    what was published.
    assert lines["year.one"].startswith("count=1|payload=2026|"), lines["year.one"]
    assert "|verdict=accepted|" in lines["year.one"]
    assert lines["year.three"].startswith("count=3|payload=2026,2027,2028|"), lines["year.three"]
    assert "|verdict=accepted|" in lines["year.three"]
    assert lines["year.ten"].startswith("count=10|"), lines["year.ten"]
    assert "|verdict=accepted|" in lines["year.ten"]
    assert lines["year.whole-window-published"].startswith("count=200|"), lines["year.whole-window-published"]
    assert "|verdict=accepted|" in lines["year.whole-window-published"]
    # 5. 200 CATEGORIES WHILE TEN YEARS ARE PUBLISHED.
    assert "|verdict=refused|" in lines["year.whole-window-refused"]
    assert "200 categories where 10 are expected" in lines["year.whole-window-refused"]
    # 6. WRONG FIRST, WRONG LAST, WRONG ORDER, WRONG COUNT, FABRICATED.
    assert "category 1 is 2025, expected 2026" in lines["year.wrong-first"]
    assert "category 3 is 2099, expected 2028" in lines["year.wrong-last"]
    assert "category 1 is 2028, expected 2026" in lines["year.wrong-order"]
    assert "presents 2 categories where 3 are expected" in lines["year.too-few"]
    assert "presents 4 categories where 3 are expected" in lines["year.too-many"]
    assert "category 1 is 2026, expected <blank>" in lines["year.fabricated-when-empty"]
    for case in ("year.wrong-first", "year.wrong-last", "year.wrong-order",
                 "year.too-few", "year.too-many", "year.fabricated-when-empty"):
        assert "|verdict=refused|" in lines[case], (case, lines[case])
    # THE COM SHAPES ALL NORMALISE TO THE SAME ORDERED PAYLOAD.
    for case in ("shape.rectangle", "shape.plain-array", "shape.numeric-text"):
        assert lines[case].startswith("count=3|payload=2026,2027,2028|"), (case, lines[case])
        assert "|verdict=accepted|" in lines[case], (case, lines[case])
    # 7. THE LITERAL CHARTS, judged against their declared cells and nothing else.
    for case in ("literal.histogram.correct", "literal.tornado.correct"):
        assert "|verdict=accepted|" in lines[case], (case, lines[case])
    assert "presents 18 categories where 20 are expected" in lines["literal.histogram.short"]
    assert "category 1 is Labour rate, expected Steel price" in lines["literal.tornado.reordered"]
    assert "category 3 is Invented driver" in lines["literal.tornado.fabricated"]
    for case in ("literal.histogram.short", "literal.tornado.reordered",
                 "literal.tornado.fabricated"):
        assert "|verdict=refused|" in lines[case], (case, lines[case])


def test_05c_the_windows_gate_takes_its_verdict_from_xvalues_and_not_from_the_formula() -> None:
    """THE GATE'S SHAPE, AND WHAT IT IS ALLOWED TO DECIDE ON. It reports the
    formula, the XValues count, the XValues payload and the point count for
    every series; it decides on the XValues alone. The parsed SERIES category
    argument is gone from the script entirely, so a blank there can no longer
    fail a chart whose categories are present - which is the defect that cost
    the histogram and the tornado a false failure."""
    code = INSPECTOR.read_text(encoding="utf-8")
    # THE ORACLE, AND ONLY IT.
    assert "$xvalues = ConvertTo-XValueList -Raw $series.XValues" in code
    assert code.count("Compare-Payload -Actual $xvalues -Expected") == 2
    # THE FORMULA IS A DIAGNOSTIC: printed, never compared, never a failure.
    assert "The series formula, for the record only. It is not the oracle." in code
    assert code.count(".formula=' + (Get-SeriesFormula -Series $series)") == 1
    for gone in ("Split-SeriesFormula", "Get-SeriesField", "Get-CategoryRows",
                 "Test-CellRange", "BLANK XValues/categories", "$categories = Get-SeriesField"):
        assert gone not in code, f"the gate still judges the formula: {gone}"
    # THE FOUR REPORTED FIELDS, one line each per series.
    for field in (".formula=", ".xvalues.count=", ".xvalues=", ".points="):
        assert code.count(field) >= 1, field
    # THE TWO CONTRACTS, AND THE ONE THAT IS DERIVED FROM THE WORKBOOK.
    runtime = _projection()["bridge"]["annual"]["runtime_category"]
    assert f"$script:CategoryWindowName = '{runtime['window_name']}'" in code
    assert f"$script:YearCountName = '{runtime['extent_name']}'" in code
    assert "$slice = $windowRange.Resize($take, 1)" in code
    assert "$script:ExpectedYear = ConvertTo-CellList -Target $slice" in code
    assert "presents the whole reserved year window" in code
    assert "presents different categories from series 1" in code
    # THE LITERAL EXPECTATIONS COME FROM THE ACCEPTED PROJECTION, so the gate
    # spells no address of its own.
    assert "$reference = [string]$chart.categories.range" in code
    assert "$script:ExpectedLiteral[$title] = ConvertTo-CellList -Target $sourceRange" in code
    assert re.search(r'"\$?[A-Z]{1,3}\$?\d+', code) is None, "the gate spells a cell address"
    # THE WORKBOOK-SCOPED NAMES ARE READ FROM THE WORKBOOK, not from one sheet.
    assert "$names = $workbook.Names" in code
    assert "$results.Names" not in code
    # EVENTS ON, because the binding is applied when the workbook opens.
    assert "$excel.EnableEvents = $true" in code
    assert "$workbooks.Open($resolved, 0, $true)" in code
    # STILL READ-ONLY, STILL A GATE.
    assert "$workbook.Close($false)" in code
    assert "if ($script:Failures.Count -eq 0) { exit 0 } else { exit 1 }" in code
    assert "CHART BINDING FAIL" in code and "CHART BINDING PASS" in code
    for banned in (".Value2 =", "PCCM_", "SaveAs", ".Save()", "ListRows.Add"):
        assert banned not in code, banned
    # AND IT NAMES THE FOUR CHARTS THE PROJECTION NAMES.
    titles = {chart["key"]: chart["title"] for chart in _projection()["charts"]}
    for key in ("s_curve", "annual_cash_flow", "histogram", "tornado"):
        assert f"'{titles[key]}'" in code, key
    assert "chartAnnual_calendar_year" not in code, "the gate still wants the failed name"

def test_10_the_annual_cash_flow_is_still_a_column_chart_of_the_published_profile() -> None:
    spec = _by_key()["annual_cash_flow"]
    assert spec["title"] == "Annual Cash Flow" and spec["kind"] == "column"
    chart = _chart_objects()["Annual Cash Flow"]
    assert type(chart).__name__ == "BarChart" and chart.type == "col"
    was = {c["key"]: c for c in _previous_charts()["charts"]}["annual_cash_flow"]
    assert [(s["key"], s["name"]) for s in spec["series"]] == [
        (s["key"], s["name"]) for s in was["series"]] == [("annual_nominal", "Annual Nominal")]
    assert spec["categories"]["key"] == was["categories"] == "calendar_year"
    # ITS CATEGORY IS BOUND AT RUNTIME; only its value carries a name.
    assert spec["categories"]["binding"] is None
    assert spec["categories"]["runtime_binding"] == "chartAnnual_category_window"
    assert spec["series"][0]["binding"] == "chartAnnual_annual_nominal"
    assert spec["value_axis_scale"] is None and spec["category_label_interval"] is None


# ===========================================================================
# 3. TOP DRIVERS BY RANK CORRELATION
# ===========================================================================
def test_20_the_tornado_value_axis_declares_a_step_and_no_bounds() -> None:
    """THE SAFETY CORRECTION. The tornado RANKS by |rho| and PLOTS signed rho,
    so a top-ranked negative driver is drawn to the LEFT of zero and a strong
    positive one can exceed any figure chosen at build time. The earlier fixed
    0..0.45 axis would have hidden the first and clipped the second - a
    presentation setting deciding what a reader sees of a published analytical
    value. The step and the format stay; the bounds are Excel's, read off the
    plotted data.

    THE EQUIVALENT REPRESENTATION, STATED. openpyxl calls the VALUE axis
    `y_axis` for every chart type, and on a horizontal bar chart Excel draws
    that axis along the bottom - so the "x-axis" step a reader sees is
    `<c:valAx><c:majorUnit/>` in the chart part, and the absence of bounds is
    `<c:scaling>` carrying its orientation and nothing else."""
    spec = _by_key()["tornado"]
    assert spec["kind"] == "bar"
    # 1. NO MINIMUM AND NO MAXIMUM, anywhere: the manifest, the projection, the
    #    loaded chart and the chart part.
    assert spec["value_axis_scale"] == {"major_unit": 0.1}
    assert "min" not in spec["value_axis_scale"] and "max" not in spec["value_axis_scale"]
    declared = _shell()["charts"]["charts"]
    tornado = next(c for c in declared if str(c["key"]) == "tornado")
    assert tornado["value_axis_scale"] == {"major_unit": 0.10}
    chart = _chart_objects()["Top Drivers by Rank Correlation"]
    assert chart.y_axis.scaling.min is None and chart.y_axis.scaling.max is None
    part = _part_for("Top Drivers by Rank Correlation")
    value_axis = part[part.index("<valAx>"): part.index("</valAx>")]
    scaling = re.search(r"<scaling>.*?</scaling>", value_axis, re.S).group(0)
    assert scaling == "<scaling><orientation val=\"minMax\" /></scaling>", scaling
    assert re.search(r"<min val=", value_axis) is None, value_axis
    assert re.search(r"<max val=", value_axis) is None, value_axis
    # 2. THE STEP REMAINS 0.10, and 3. THE FORMAT REMAINS 0.00.
    assert chart.y_axis.majorUnit == 0.1
    assert '<majorUnit val="0.1"' in value_axis, value_axis
    assert spec["value_axis_format"] == "0.00"
    assert chart.y_axis.numFmt.formatCode == "0.00"
    assert chart.y_axis.numFmt.sourceLinked is False
    assert 'formatCode="0.00"' in value_axis, value_axis
    # AND NO OTHER CHART CARRIES A BOUND EITHER: a value a reader cannot see is
    # not a presentation choice this workbook makes anywhere.
    for key, other in _by_key().items():
        scale = other["value_axis_scale"]
        assert scale is None or set(scale) == {"major_unit"}, (key, scale)
        plot = _chart_objects()[other["title"]]
        for axis in (plot.y_axis, plot.x_axis):
            assert axis.scaling.min is None and axis.scaling.max is None, key
    for chart_part in _chart_parts():
        axes = re.findall(r"<scaling>.*?</scaling>", chart_part, re.S)
        assert axes and all(a == "<scaling><orientation val=\"minMax\" /></scaling>"
                            for a in axes), axes


def test_20b_a_bound_on_any_chart_value_axis_is_refused() -> None:
    """THE REGRESSION THIS CORRECTION FORBIDS: a minimum or a maximum coming
    back, as the projection and as the manifest. Both layers refuse it by name,
    so the safety property is structural rather than merely absent."""
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    from pccm_builder import load_sim_contract
    inspection = build_phase8_charts_inspection(
        load_spec(MANIFEST), structure.limits.max_generated_year_columns,
        load_sim_contract(SPEC / "sim_contract.yaml"))
    validate_phase8_charts_inspection(inspection)
    assert inspection == _projection()
    for bound in ({"min": 0.0}, {"max": 0.45}, {"min": 0.0, "max": 0.45}):
        damaged = copy.deepcopy(inspection)
        chart = next(c for c in damaged["charts"] if c["key"] == "tornado")
        chart["value_axis_scale"] = dict(chart["value_axis_scale"], **bound)
        with pytest.raises(ValueError, match="bound can hide a plotted value"):
            validate_phase8_charts_inspection(damaged)
    damaged = copy.deepcopy(inspection)
    chart = next(c for c in damaged["charts"] if c["key"] == "tornado")
    chart["value_axis_scale"] = {"major_unit": 0.0}
    with pytest.raises(ValueError, match="is not a step"):
        validate_phase8_charts_inspection(damaged)
    # AND THE MANIFEST REFUSES IT BEFORE THE PROJECTION IS EVER BUILT.
    text = MANIFEST.read_text(encoding="utf-8")
    old = "        value_axis_scale:\n          major_unit: 0.10\n"
    assert text.count(old) == 1
    for injected, message in (
        ("        value_axis_scale:\n          min: 0\n          major_unit: 0.10\n",
         "value_axis_scale declares ['min']"),
        ("        value_axis_scale:\n          max: 0.45\n          major_unit: 0.10\n",
         "value_axis_scale declares ['max']"),
    ):
        path = MANIFEST.with_name("workbook.bounds-mutation.yaml")
        path.write_text(text.replace(old, injected, 1), encoding="utf-8")
        try:
            with pytest.raises(SpecError, match=re.escape(message)):
                load_spec(path)
        finally:
            path.unlink(missing_ok=True)


def test_20c_no_plotted_rho_can_be_hidden_or_clipped_by_the_chart() -> None:
    """4 AND 5, STATED AS THE PROPERTY THEY ARE. The bridge mirrors the SIGNED
    rho the Sensitivity sheet publishes, cell for cell, with no absolute value,
    no clamp, no floor and no ceiling; the chart declares no bound; so a
    negative rho and a rho above any particular figure both reach the plot
    exactly as published."""
    block = _projection()["bridge"]["drivers"]
    results = _workbook()[_projection()["bridge_sheet"]]
    rho = next(c for c in block["columns"] if c["key"] == "rho")
    source = {c["key"]: c["column"] for c in _projection()["sensitivity_source"]["columns"]}
    first_source = int(_projection()["sensitivity_source"]["first_row"])
    sheet = _projection()["sensitivity_sheet"]
    eligibility = _projection()["sensitivity_source"]["eligibility"]["column"]
    for offset in range(int(block["row_count"])):
        formula = results[f"{rho['column']}{int(block['first_row']) + offset}"].value
        row = first_source + offset
        # THE MIRROR, EXACTLY: eligible and non-blank -> the published cell itself.
        assert formula == (
            f"=IF({sheet}!${eligibility}${row}=\"\",NA(),"
            f"IF({sheet}!${source['rho']}${row}=\"\",NA(),{sheet}!${source['rho']}${row}))"), formula
        for banned in ("ABS(", "MAX(", "MIN(", "IFERROR(", "ROUND(", "-"):
            assert banned not in formula.replace("=IF(", "", 1) or banned == "-", (banned, formula)
        assert "ABS(" not in formula and "MAX(" not in formula and "MIN(" not in formula
    # THE SIGN IS THE POINT OF THE CHART, and the manifest says so.
    manifest = MANIFEST.read_text(encoding="utf-8")
    assert "ranked by the ABSOLUTE value of rho" in manifest
    assert "a strongly negative driver ranks alongside an equally strong positive one" in manifest
    # AND THE SOURCE IS THE SIGNED COLUMN, never the magnitude beside it.
    assert not any(c["key"].startswith("abs_")
                   for c in _projection()["sensitivity_source"]["columns"])
    assert [c["key"] for c in block["columns"]] == ["driver_name", "rho"]


def test_20d_the_safety_correction_moved_the_bounds_and_nothing_else() -> None:
    """6 AND 7 TOGETHER. The whole chart projection is rebuilt from the manifest
    as it stood at e87d566 - the accepted chart package - and compared with the
    one this tree produces, field for field. The tornado's value-axis scale is
    the ONLY difference: the ranking source and its ranges, both applied-year
    bindings, the histogram's label interval and the s-curve's title all come
    through identical."""
    import tempfile
    import yaml as _yaml
    from pccm_builder import load_sim_contract
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    window = structure.limits.max_generated_year_columns
    accepted_text = subprocess.run(
        ["git", "show", f"{ACCEPTED_CHART_PACKAGE}:pccm/spec/workbook.yaml"],
        cwd=PCCM_ROOT.parent, check=True, stdout=subprocess.PIPE).stdout.decode("utf-8")
    scratch = Path(tempfile.mkdtemp(prefix="pccm-bounds-")) / "workbook.yaml"
    scratch.write_text(accepted_text, encoding="utf-8")
    # THE ACCEPTED PACKAGE DID CARRY THE BOUNDS - and the loader now REFUSES it
    # unread, which is the correction's whole point and is stronger than any
    # comparison: the manifest that produced the unsafe axis can no longer be
    # built at all.
    accepted_chart = next(
        c for c in _yaml.safe_load(accepted_text)["phase6_shell"]["charts"]["charts"]
        if str(c["key"]) == "tornado")
    assert accepted_chart["value_axis_scale"] == {"min": 0, "max": 0.45, "major_unit": 0.10}
    with pytest.raises(SpecError, match=re.escape("value_axis_scale declares ['max', 'min']")):
        load_spec(scratch)
    # SO THE COMPARISON IS MADE ACROSS THE DECLARED DELTA AND NOTHING ELSE: the
    # two bound lines taken out of the accepted manifest, everything else its
    # own bytes. If any other byte of the chart layer had moved, the projection
    # below would differ somewhere other than the scale - and it does not.
    without_bounds = accepted_text.replace(
        "        value_axis_scale:\n          min: 0\n          max: 0.45\n          major_unit: 0.10\n",
        "        value_axis_scale:\n          major_unit: 0.10\n", 1)
    assert without_bounds != accepted_text
    scratch.write_text(without_bounds, encoding="utf-8")
    before = build_phase8_charts_inspection(load_spec(scratch), window,
                                            load_sim_contract(SPEC / "sim_contract.yaml"))
    now = _projection()
    assert set(before) == set(now)
    # THE BRIDGE GAINED THE RUNTIME CATEGORY CONTRACT and the category column
    # lost its cut name; everything else about the layer is unchanged.
    for key in set(before) - {"charts", "bridge"}:
        assert before[key] == now[key], f"{key} moved"
    was_bridge = copy.deepcopy(before["bridge"])
    now_bridge = copy.deepcopy(now["bridge"])
    runtime = now_bridge["annual"].pop("runtime_category")
    assert now_bridge["annual"]["applied_binding"].pop("runtime_category") == runtime["key"]
    for block in (was_bridge["annual"], now_bridge["annual"]):
        for column in block["columns"]:
            if column["key"] == runtime["key"]:
                column["applied_binding"] = None
    assert was_bridge == now_bridge, "the bridge moved beyond the runtime category"
    was = {c["key"]: c for c in before["charts"]}
    current = {c["key"]: c for c in now["charts"]}
    assert list(was) == list(current), "the chart order moved"
    for key, chart in was.items():
        after = current[key]
        assert set(chart) == set(after), key
        for field in set(chart) - {"categories", "series"}:
            assert chart[field] == after[field], (key, field)
        # THE CATEGORY AND THE SERIES CARRY THE RUNTIME CONTRACT NOW; their
        # KEYS, names and the window each is written over are unchanged.
        assert chart["categories"]["key"] == after["categories"]["key"], key
        assert chart["categories"]["window_range"] == after["categories"]["window_range"], key
        assert [(s["key"], s["name"], s["window_range"]) for s in chart["series"]] == \
            [(s["key"], s["name"], s["window_range"]) for s in after["series"]], key
        if key == "tornado":
            assert after["value_axis_scale"] == {"major_unit": 0.1}
            assert after["categories"] == chart["categories"], key
    # AND THE THREE OTHER CORRECTIONS, NAMED, straight out of that comparison.
    assert current["s_curve"]["title"] == was["s_curve"]["title"] == "Cumulative Cost Profile"
    for year_chart in YEAR_CHARTS:
        assert current[year_chart]["categories"]["binding"] is None
        assert current[year_chart]["categories"]["runtime_binding"] == "chartAnnual_category_window"
        assert [s["binding"] for s in current[year_chart]["series"]] == \
            [s["binding"] for s in was[year_chart]["series"]]
    assert current["histogram"]["category_label_interval"] == 2
    assert _chart_objects()["Total Cost Distribution"].x_axis.tickLblSkip == 2
    assert current["tornado"]["categories"]["range"] == was["tornado"]["categories"]["range"]
    assert [s["range"] for s in current["tornado"]["series"]] == \
        [s["range"] for s in was["tornado"]["series"]]


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
        "builder/pccm_builder/phase8_charts.py", "builder/pccm_builder/spec_loader.py",
        # The structural verification, which derives the applied-year chart
        # names from the manifest now that they are workbook-scoped.
        "builder/pccm_builder/verify.py",
        # The presentation owner's declaration, and the three owners that reach
        # a presentation boundary and call it. modChartPresentation.bas is a NEW
        # file and carries no reversal: it did not exist at ebeae65.
        "spec/structure_contract.yaml", "src/vba/ThisWorkbook.vba",
        "src/vba/modReset.bas", "src/vba/modSimAnnualRun.bas"}
    for name in DECLARED_CHART_POLISH_CHANGES:
        now = (PCCM_ROOT / name).read_text(encoding="utf-8")
        then = _git_show(name)
        assert now != then, name
        assert strip_chart_polish(name, now) == then, name
        assert strip_chart_polish(name, then) == then, name
        # A LAYER THAT IS ONLY PARTLY THERE IS REFUSED, not quietly half-taken
        # off: one declared fragment of this file's own layer, edited.
        from chart_polish_declaration import _HUNKS
        if len(_HUNKS[name]) < 2:
            # ONE FRAGMENT IS THE WHOLE LAYER for this file: removing it leaves
            # nothing partly present to refuse, which is the reversal's own
            # all-or-none rule and is covered by the round trip above.
            continue
        first = _HUNKS[name][0][0]
        damaged = now.replace(first, "# an undeclared edit\n", 1)
        assert damaged != now and first not in damaged, name
        with pytest.raises(AssertionError, match="partly present"):
            strip_chart_polish(name, damaged)


def test_41_production_moved_only_by_the_declared_presentation_owner() -> None:
    """NO ANALYTICAL OWNER MOVED. The calculation, the simulation, the
    fingerprint, the statistics, the annual store and the state derivations are
    byte-identical to ebeae65. What changed is the new presentation module, the
    structure contract entry that declares it, and one call line in each of the
    three owners that reach a presentation boundary - the workbook opening, an
    annual run that published, a reset that cleared."""
    changed = {line for line in subprocess.run(
        ["git", "diff", "--name-only", PREVIOUS_HEAD, "--", "pccm/src", "pccm/spec"],
        cwd=PCCM_ROOT.parent, check=True, stdout=subprocess.PIPE).stdout.decode().split()}
    # A MODULE ADDED BUT NOT YET COMMITTED IS STILL A CHANGE TO PRODUCTION.
    changed |= {line for line in subprocess.run(
        ["git", "ls-files", "--others", "--exclude-standard", "pccm/src", "pccm/spec"],
        cwd=PCCM_ROOT.parent, check=True, stdout=subprocess.PIPE).stdout.decode().split()}
    assert changed == {
        "pccm/spec/structure_contract.yaml",
        "pccm/spec/workbook.yaml",
        "pccm/src/vba/ThisWorkbook.vba",
        "pccm/src/vba/modChartPresentation.bas",
        "pccm/src/vba/modReset.bas",
        "pccm/src/vba/modSimAnnualRun.bas",
    }, sorted(changed)
    # EVERY OTHER CONTRACT IS UNTOUCHED.
    for path in ("spec/sim_contract.yaml", "spec/calc_contract.yaml",
                 "spec/input_contract.yaml", "spec/driver_contract.yaml"):
        done = subprocess.run(["git", "diff", "--quiet", PREVIOUS_HEAD, "--", f"pccm/{path}"],
                              cwd=PCCM_ROOT.parent)
        assert done.returncode == 0, f"{path} changed since {PREVIOUS_HEAD}"
    # AND THE THREE TOUCHED OWNERS GAINED A CALL AND LOST NOTHING.
    for module in ("ThisWorkbook.vba", "modReset.bas", "modSimAnnualRun.bas"):
        diff = subprocess.run(["git", "diff", PREVIOUS_HEAD, "--", f"pccm/src/vba/{module}"],
                              cwd=PCCM_ROOT.parent, check=True,
                              stdout=subprocess.PIPE).stdout.decode()
        removed = [line for line in diff.splitlines()
                   if line.startswith("-") and not line.startswith("---")]
        assert removed == [], (module, removed)
        added = "\n".join(line for line in diff.splitlines()
                           if line.startswith("+") and not line.startswith("+++"))
        assert "modChartPresentation.ChartPresentationApplyCategories(chartDetail)" in added, module
        for banned in ("Value2", "SimAnnualStore", "CalcReport", "SimReport", "Fingerprint"):
            assert banned not in added, (module, banned)


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
