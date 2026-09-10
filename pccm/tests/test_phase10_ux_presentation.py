#!/usr/bin/env python3
"""P10-UX: the first corrections found by looking at the real workbook.

WHY THIS FILE EXISTS SEPARATELY FROM THE PHASE-8 CHART SUITE. That suite proves
what the charts PLOT - the ranges, the identities, the state each one inherits,
the semantics of an absent point. None of it could have caught what a human
found by opening the file: that the plots were too small to read, that a button
was drawn on top of the paragraph explaining the block it sat in, and that a
sheet with one visible column told nobody where to type. Those are presentation
facts, and until now nothing asserted any of them.

THE THREE CORRECTIONS, AND WHAT EACH ONE IS NOT.

  UX-001  The charts are two and a half times the plot area they were, the axes
          carry a compact number format and a smaller label font, the tornado's
          driver names are pinned to the low end of its value axis, and an
          absent CATEGORY is blank instead of #N/A. Not one source range, series
          identity, Top-N rule, eligibility gate or state word moves, and an
          absent VALUE is still NA() - which is what keeps "no result" from
          being drawn as a project costing nothing.

  UX-002  The Apply / Update Timeline button moves out of the Applied Timeline
          block, where it was drawn across the explanatory note, and into the
          Setup command area as the first command. Same endpoint, same caption,
          same semantics; a different anchor.

  UX-003  The Inflation sheet's no-timeline message becomes an instruction and
          is rendered as one. It is the SAME formula over the SAME applied
          timeline owner: no new state, no new status, no second button.

Runs standalone or under pytest.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

import pytest  # noqa: E402
import yaml  # noqa: E402

from pccm_builder import load_structure_contract  # noqa: E402

SPEC = PCCM_ROOT / "spec"
BUILD = PCCM_ROOT / "build"
SRC = PCCM_ROOT / "src" / "vba"
WORKBOOK = BUILD / "PCCM_stageA.xlsx"
CHART_PROJECTION = BUILD / "phase8_charts_inspection.json"

# THE GEOMETRY THE MANUAL REVIEW WAS PERFORMED AGAINST. Every chart was this
# size, and the reviewer found them cramped while the workbook was still EMPTY.
# Kept as a number rather than as a memory so "materially larger" is a measured
# claim and not an opinion.
REVIEWED_WIDTH_CM = 9.6
REVIEWED_HEIGHT_CM = 7.2
REVIEWED_AREA_CM2 = REVIEWED_WIDTH_CM * REVIEWED_HEIGHT_CM

# THE FOUR PLOTS, BY KEY. Named so a fifth appearing - or one of these being
# quietly dropped while the others grow - fails here.
CHART_KEYS = ("s_curve", "histogram", "annual_cash_flow", "tornado")

# WHERE THE TIMELINE BUTTON WAS, AND WHY IT WAS WRONG. E43 is the "Applied
# Timeline" section heading row and E is that block's NOTE column.
REVIEWED_TIMELINE_ANCHOR = "E43"
TIMELINE_ENTRY_POINT = "PCCM_ApplyTimeline"
TIMELINE_SHAPE = "btnPCCMApplyTimeline"

_CACHE: dict[str, object] = {}


def _shell() -> dict:
    if "shell" not in _CACHE:
        _CACHE["shell"] = yaml.safe_load(
            (SPEC / "workbook.yaml").read_text(encoding="utf-8"))["phase6_shell"]
    return _CACHE["shell"]  # type: ignore[return-value]


def _charts() -> dict:
    return _shell()["charts"]


def _projection() -> dict:
    if "projection" not in _CACHE:
        _CACHE["projection"] = json.loads(CHART_PROJECTION.read_text(encoding="utf-8"))
    return _CACHE["projection"]  # type: ignore[return-value]


def _by_key() -> dict:
    return {chart["key"]: chart for chart in _projection()["charts"]}


def _structure():
    if "structure" not in _CACHE:
        _CACHE["structure"] = load_structure_contract(SPEC / "structure_contract.yaml")
    return _CACHE["structure"]


def _workbook():
    if "workbook" not in _CACHE:
        import openpyxl
        _CACHE["workbook"] = openpyxl.load_workbook(WORKBOOK)
    return _CACHE["workbook"]


def _column_index(letter: str) -> int:
    total = 0
    for char in letter:
        total = total * 26 + (ord(char) - 64)
    return total


def _anchor(cell: str) -> tuple[str, int]:
    match = re.match(r"([A-Z]+)(\d+)$", cell)
    assert match, cell
    return match.group(1), int(match.group(2))


# ===========================================================================
# UX-001. THE CHARTS
# ===========================================================================
def test_01_every_chart_is_materially_larger_than_the_reviewed_geometry() -> None:
    """1. Materially larger, measured against the size the review was done at.

    Half again in each direction is the floor, not the target: the charts are
    16.0 x 11.5 cm against 9.6 x 7.2, which is two and a half times the area.
    The floor is what this asserts, so a later change that shrinks them back
    towards the reviewed size fails without anybody having to remember why.
    """
    for key in CHART_KEYS:
        chart = _by_key()[key]
        assert chart["width_cm"] >= REVIEWED_WIDTH_CM * 1.5, (key, chart["width_cm"])
        assert chart["height_cm"] >= REVIEWED_HEIGHT_CM * 1.5, (key, chart["height_cm"])
        area = chart["width_cm"] * chart["height_cm"]
        assert area >= REVIEWED_AREA_CM2 * 2.0, (key, area)


def test_02_there_are_still_exactly_four_charts_and_they_are_the_same_four() -> None:
    """2. Four chart objects, by key and by title. Enlarging a Dashboard is not
    a licence to drop the plot that would not fit."""
    assert [chart["key"] for chart in _projection()["charts"]] == list(CHART_KEYS)
    assert len(_workbook()["Dashboard"]._charts) == len(CHART_KEYS)
    titles = {chart["key"]: chart["title"] for chart in _projection()["charts"]}
    assert titles == {
        "s_curve": "Cumulative Cost",
        "histogram": "Total Cost Distribution",
        "annual_cash_flow": "Annual Cash Flow",
        "tornado": "Top Drivers by Rank Correlation",
    }, titles


def test_03_no_chart_source_identity_moved() -> None:
    """3. THE PRESENTATION BATCH TOUCHES NO DATA. Every series range, category
    range, source block, authority and state qualifier is compared against the
    accepted Reset-Results tree - the last commit before a human looked at the
    workbook - field for field.
    """
    import subprocess
    accepted = json.loads(subprocess.run(
        ["git", "show", "a7c2222:pccm/build/phase8_charts_inspection.json"],
        cwd=PCCM_ROOT.parent, stdout=subprocess.PIPE).stdout.decode() or "{}")
    if not accepted:
        # The projection is a build artefact and is not tracked, so the accepted
        # copy is reconstructed from the accepted MANIFEST instead - which is the
        # authority the projection is derived from either way.
        manifest = yaml.safe_load(subprocess.run(
            ["git", "show", "a7c2222:pccm/spec/workbook.yaml"],
            cwd=PCCM_ROOT.parent, check=True,
            stdout=subprocess.PIPE).stdout.decode())["phase6_shell"]["charts"]
        was = {str(chart["key"]): chart for chart in manifest["charts"]}
        now = {str(chart["key"]): chart for chart in _charts()["charts"]}
        assert set(was) == set(now)
        for key, before in was.items():
            after = now[key]
            for field in ("kind", "source", "categories", "state_source",
                          "also_qualified_by"):
                assert before.get(field) == after.get(field), (key, field)
            assert before["series"] == after["series"], key
            assert before["title"] == after["title"], key
        # AND THE BRIDGE BLOCKS THEMSELVES: same rows, same columns, same
        # letters. Only the `absent` marker on the three CATEGORY columns is
        # allowed to differ, and test_04 is what checks it.
        was_bridge = manifest["bridge"]
        now_bridge = _charts()["bridge"]
        for block in ("annual", "distribution", "drivers"):
            for field in ("heading_row", "header_row", "first_row"):
                assert was_bridge[block][field] == now_bridge[block][field], block
            for before_col, after_col in zip(was_bridge[block]["columns"],
                                             now_bridge[block]["columns"]):
                for field in ("key", "header", "column", "format"):
                    assert before_col.get(field) == after_col.get(field), (block, field)


def test_04_an_absent_category_is_blank_and_an_absent_value_is_still_na() -> None:
    """4. NO VISIBLE #N/A, AND NOTHING CONVERTED TO ZERO.

    The distinction is the whole correction: an absent CATEGORY is a missing
    LABEL and is blank, an absent VALUE is a missing POINT and is still NA().
    Excel draws no bar, no marker and no column for NA(); it draws the literal
    text "#N/A" for a category. This asserts both halves for every chart, and
    asserts per COLUMN that no value column was blanked along with the label.
    """
    for key in CHART_KEYS:
        chart = _by_key()[key]
        assert chart["no_data_category"] == '""', key
        assert chart["no_data_value"] == "NA()", key
    for name in ("annual", "distribution", "drivers"):
        block = _projection()["bridge"][name]
        categories = {chart["categories"]["key"] for chart in _projection()["charts"]
                      if chart["source_block"] == name}
        for column in block["columns"]:
            expected = '""' if column["key"] in categories else "NA()"
            assert column["absent"] == expected, (name, column["key"], column["absent"])
    # AND THE FORMULAS ON THE SHEET SAY THE SAME THING. A projection that
    # described a workbook the builder did not write would be worse than none.
    results = _workbook()["Results"]
    annual = _charts()["bridge"]["annual"]
    columns = {str(c["key"]): str(c["column"]) for c in annual["columns"]}
    row = int(annual["first_row"])
    assert '""' in results[f"{columns['calendar_year']}{row}"].value
    assert "NA()" not in results[f"{columns['calendar_year']}{row}"].value
    assert ",NA()," in results[f"{columns['annual_nominal']}{row}"].value
    # NOT ZERO. Neither cell may fall back to a number.
    for key in ("calendar_year", "annual_nominal"):
        assert ",0," not in results[f"{columns[key]}{row}"].value, key


def test_05_the_tornado_has_the_most_category_room_and_its_labels_are_pinned() -> None:
    """5. TWO SEPARATE THINGS, AND THE SECOND MATTERS MORE.

    Rho is SIGNED, so the tornado's value axis crosses at zero in the MIDDLE of
    the plot, and Excel's default puts each driver's name beside that crossing -
    over the bars, clipped by them, and moving as the data changes. Pinning the
    category labels to the low end is what actually gives a long driver
    description somewhere to be.

    RESTATED AT UX-004. This used to say the tornado is WIDER than its
    neighbours, which is how UX-001 bought that room and is what made the
    Dashboard look unbalanced in real Excel. The room is the same claim; where
    it comes from is not. It is now an INTERNAL allocation, asserted in
    centimetres, and the outer width is required to match the other three
    exactly - which test_14 is what states.
    """
    tornado = _by_key()["tornado"]
    assert tornado["kind"] == "bar", "a tornado is a horizontal bar chart"
    others = [chart["width_cm"] for key, chart in _by_key().items() if key != "tornado"]
    assert tornado["width_cm"] == max(others), (tornado["width_cm"], others)
    assert _projection()["axis_presentation"]["category_label_position"] == "low"
    # AND IT REACHES THE WORKBOOK. openpyxl calls the CATEGORY axis `x_axis` for
    # every chart type here, including the horizontal bar Excel draws the other
    # way round, so this is the axis the driver names are on.
    charts = {chart.title.tx.rich.p[0].r[0].t: chart
              for chart in _workbook()["Dashboard"]._charts}
    plot = charts["Top Drivers by Rank Correlation"]
    assert plot.x_axis.tickLblPos == "low"
    assert plot.y_axis.numFmt.formatCode == _charts()["number_formats"]["rho"]


def test_06_the_axes_carry_a_compact_format_and_a_smaller_label_font() -> None:
    """2. Density, by the two levers a build can actually pull.

    A COMPACT FORMAT THAT CHANGES NO VALUE. Excel's scaling comma divides what
    is displayed, never what is stored - so a money axis reads "1,200M" and the
    cell behind it is still 1,200,000,000. Three conditional sections rather
    than one fixed scale, because a single ,, prints a half-million-riyal
    project as "0M", and a compact format that lies about the order of
    magnitude is worse than a wide one that does not fit.

    AND NO BUILD-TIME TICK INTERVAL. A skip chosen here would have to be chosen
    against the 200-column ceiling and would hide four labels in five on a
    ten-year project. Excel's automatic interval reads the rendered width at
    open time; enlarging the plot and shrinking the label is how this batch
    makes that automatic choice a good one.
    """
    formats = _charts()["number_formats"]
    compact = formats["money_axis"]
    assert compact.count(";") == 2, compact
    assert '"M"' in compact and '"k"' in compact, compact
    assert ",," in compact, "the millions section does not scale"
    for key in ("s_curve", "annual_cash_flow"):
        assert _by_key()[key]["value_axis_format"] == compact, key
    assert _by_key()["histogram"]["category_axis_format"] == compact
    size = _projection()["axis_presentation"]["label_font_size"]
    assert 700 <= size < 1000, size
    for chart in _workbook()["Dashboard"]._charts:
        for axis in (chart.x_axis, chart.y_axis):
            assert axis.txPr is not None, "an axis carries no declared label font"
            assert axis.txPr.p[0].pPr.defRPr.sz == size
            assert axis.numFmt is not None
            assert axis.numFmt.sourceLinked is False, (
                "a format Excel is free to ignore is not a format")
    # THE BLANK POLICY IS READ FROM THE FILE ITSELF. openpyxl writes
    # dispBlanksAs into the chart part but does not read it back onto the loaded
    # object, so the loaded workbook is the wrong place to ask - and asking the
    # projection alone would only prove the builder agrees with itself.
    declared = _projection()["axis_presentation"]["display_blanks_as"]
    for part in _chart_parts():
        assert f'<dispBlanksAs val="{declared}"' in part, part[:80]
        # AND THE COMPACT FORMAT SURVIVED THE ROUND TRIP INTO THE FILE.
        assert "sourceLinked=\"0\"" in part


def test_07_the_enlarged_charts_do_not_overlap_and_stay_inside_the_region() -> None:
    """3 and 4. Bigger is only better if they still fit beside each other. The
    2x2 stays: two columns, two rows, and no plot reaching into its neighbour."""
    region = _shell()["dashboard"]["chart_region"]
    placed = []
    for chart in _projection()["charts"]:
        letter, row = _anchor(chart["anchor"])
        assert int(region["first_row"]) <= row <= int(region["last_row"]), chart["key"]
        placed.append((chart["key"], _column_index(letter), row,
                       chart["width_cm"], chart["height_cm"]))
    columns = sorted({column for _k, column, _r, _w, _h in placed})
    rows = sorted({row for _k, _c, row, _w, _h in placed})
    assert len(columns) == 2 and len(rows) == 2, (columns, rows)
    # THE LEFT PLOT MUST END BEFORE THE RIGHT ONE STARTS, in centimetres, using
    # the Dashboard's own declared column widths. An enlargement that quietly
    # overlapped would look like a rendering fault rather than a decision.
    widths = _dashboard_column_widths()
    left = _column_offset_cm(widths, columns[0])
    right = _column_offset_cm(widths, columns[1])
    widest_left = max(w for _k, c, _r, w, _h in placed if c == columns[0])
    assert left + widest_left <= right, (left, widest_left, right)
    # AND THE TOP ROW MUST END BEFORE THE BOTTOM ROW STARTS.
    gap_rows = rows[1] - rows[0]
    default_row_cm = float(_presentation()["row_heights"]["default"]) / 72 * 2.54
    tallest_top = max(h for _k, _c, r, _w, h in placed if r == rows[0])
    assert gap_rows * default_row_cm >= tallest_top, (gap_rows, tallest_top)


def _chart_parts() -> list[str]:
    """The four chart XML parts, read out of the built workbook.

    A chart's axis settings live in its own part, not on the sheet, and openpyxl
    round-trips only some of them onto the loaded object. Reading the parts is
    the only way to assert what Excel will actually be handed.
    """
    import zipfile
    with zipfile.ZipFile(WORKBOOK) as archive:
        names = sorted(n for n in archive.namelist()
                       if n.startswith("xl/charts/chart") and n.endswith(".xml"))
        assert len(names) == len(CHART_KEYS), names
        return [archive.read(name).decode("utf-8") for name in names]


def _presentation() -> dict:
    return yaml.safe_load(
        (SPEC / "workbook.yaml").read_text(encoding="utf-8"))["presentation"]


def _dashboard_column_widths() -> dict[int, float]:
    sheets = yaml.safe_load((SPEC / "workbook.yaml").read_text(encoding="utf-8"))["sheets"]
    dashboard = next(sheet for sheet in sheets if sheet["name"] == "Dashboard")
    return {_column_index(letter): float(width)
            for letter, width in dashboard["column_widths"].items()}


def _column_offset_cm(widths: dict[int, float], index: int) -> float:
    """Where a column's left edge sits, in centimetres from the sheet's edge.

    Excel's column "width" is a character count; the pixel width is 7n + 5 and
    a pixel is 1/96 inch. Columns the Dashboard does not declare take Excel's
    default of 8.43. This is arithmetic over the manifest, not a measurement.
    """
    default = 8.43
    total = 0.0
    for column in range(1, index):
        chars = widths.get(column, default)
        total += (7 * chars + 5) / 96 * 2.54
    return total


# ===========================================================================
# UX-002. THE TIMELINE BUTTON
# ===========================================================================
def test_08_the_timeline_button_no_longer_sits_over_the_applied_block() -> None:
    """6. It was drawn on the section heading row, in the note column, of the
    block it explains. Both facts are asserted: it is out of the block's rows,
    and it is out of the block's note column at those rows."""
    structure = _structure()
    button = structure.button_for(TIMELINE_ENTRY_POINT)
    assert button.anchor_cell != REVIEWED_TIMELINE_ANCHOR
    block = structure.applied_block
    letter, row = _anchor(button.anchor_cell)
    # EVERY ROW THE APPLIED TIMELINE BLOCK WRITES IN, heading and note included.
    occupied = {int(block["section_row"]), int(block["note_row"])}
    occupied |= {int(_anchor(field.cell)[1]) for field in structure.structural_fields}
    occupied.add(int(_anchor(structure.structural_state.cell)[1]))
    # A BUTTON IS TWO ROWS TALL AT THE DECLARED HEIGHT, so it is not enough for
    # its anchor to miss the block: its whole extent has to.
    span = _button_row_span(structure, row)
    assert not (span & occupied), (
        f"the timeline button covers applied-timeline row(s) {sorted(span & occupied)}")
    assert row > max(occupied), (
        "the button sits above the block it must not intersect")


def _button_row_span(structure, top: int) -> set[int]:
    height_points = float(structure.buttons[0].height)
    default_row = float(_presentation()["row_heights"]["default"])
    rows = max(1, int(-(-height_points // default_row)))
    return set(range(top, top + rows))


def test_09_the_timeline_button_is_the_first_command_and_still_the_same_command() -> None:
    """6 and 7. In the command area, aligned with the others, bound to the same
    endpoint, and FIRST - because applying a timeline is what generates the year
    columns every later command needs."""
    structure = _structure()
    block = structure.commands["block"]
    button = structure.button_for(TIMELINE_ENTRY_POINT)
    assert button.shape_name == TIMELINE_SHAPE
    assert button.caption == "Apply / Update Timeline"
    assert button.entry_point == TIMELINE_ENTRY_POINT
    assert button.sheet == block["sheet"]
    assert button.anchor_cell == f"{block['button_column']}{block['first_button_row']}"
    # ALIGNMENT AND SPACING ARE THE BLOCK'S, NOT THIS BUTTON'S. Every command
    # sits on the declared column at the declared pitch, in declaration order,
    # with no gap - which is what "matches the other command buttons" means when
    # it is asserted rather than eyeballed.
    commands = [b for b in structure.buttons if b.sheet == block["sheet"]]
    column = str(block["button_column"])
    first, pitch = int(block["first_button_row"]), int(block["button_row_pitch"])
    assert [b.anchor_cell for b in commands] == [
        f"{column}{first + pitch * i}" for i in range(len(commands))]
    assert commands[0].shape_name == TIMELINE_SHAPE
    # AND THE ENDPOINT IS UNTOUCHED, in the contract and in the source.
    assert TIMELINE_ENTRY_POINT in structure.entry_points
    source = "\n".join(p.read_text(encoding="utf-8") for p in sorted(SRC.glob("*.bas")))
    assert re.search(rf"^Public Sub {TIMELINE_ENTRY_POINT}\(\)", source, re.M)


def test_10_every_other_button_keeps_its_binding() -> None:
    """8. A layout batch that rebound a button would be changing what the user
    runs, not where they click. Sheet, shape, caption and endpoint of every
    button are compared against the accepted tree; only anchors may move."""
    import subprocess
    accepted = yaml.safe_load(subprocess.run(
        ["git", "show", "a7c2222:pccm/spec/structure_contract.yaml"],
        cwd=PCCM_ROOT.parent, check=True,
        stdout=subprocess.PIPE).stdout.decode())["buttons"]["definitions"]
    was = {str(b["shape_name"]): b for b in accepted}
    now = {b.shape_name: b for b in _structure().buttons}
    # P10-2C ADDED THE SIXTH COMMAND AFTER THIS BATCH, and it is named rather
    # than admitted by a loosened comparison: every button the UX batch was
    # accepted with must still be here, bound exactly as it was, and the only
    # thing that may have appeared since is Repair Profiling.
    assert set(was) - set(now) == set(), sorted(set(was) - set(now))
    assert set(now) - set(was) <= {"btnPCCMRepairProfiling"}, sorted(set(now) - set(was))
    for shape, before in was.items():
        after = now[shape]
        assert (after.sheet, after.caption, after.entry_point) == (
            str(before["sheet"]), str(before["caption"]),
            str(before["entry_point"])), shape
    # ONLY THE TIMELINE BUTTON CHANGED ROW BY MORE THAN THE ONE PITCH THE BLOCK
    # SHIFTED. The four Phase-4 register buttons did not move at all.
    for shape in ("btnPCCMAddCostLine", "btnPCCMDeleteCostLine",
                  "btnPCCMAddRisk", "btnPCCMDeleteRisk"):
        assert now[shape].anchor_cell == str(was[shape]["anchor_cell"]), shape


# ===========================================================================
# UX-003. THE INFLATION SHEET
# ===========================================================================
def test_11_inflation_states_what_to_do_when_no_timeline_is_applied() -> None:
    """9. Visible, and an INSTRUCTION rather than a fact. It names the sheet to
    go to, the command to select, and the columns that appear."""
    structure = _structure()
    message = structure.state_messages["inflation_not_applied"]
    for phrase in ("No timeline applied", "Setup", "Apply / Update Timeline",
                   "inflation-year columns"):
        assert phrase in message, phrase
    assert message in structure.state_messages["inflation_formula"]
    grid = structure.grids["inflation"]
    cell = _workbook()["Inflation"][f"B{grid.state_message_row}"]
    assert isinstance(cell.value, str) and cell.value.startswith("=")
    assert message in cell.value
    # AND IT IS RENDERED SO A USER MEETS IT. The reviewer read straight past the
    # old one because it was note-sized italic grey; this is the correction.
    presentation = _presentation()
    assert cell.font.size == presentation["sizes"]["state_message"]
    assert cell.font.size > presentation["sizes"]["note"]
    assert cell.font.bold is True
    assert cell.font.italic is not True
    assert cell.fill.fgColor.rgb.endswith(presentation["colors"]["state_message_fill"])
    worksheet = _workbook()["Inflation"]
    assert worksheet.row_dimensions[grid.state_message_row].height == float(
        presentation["row_heights"]["state_message"])


def test_12_the_message_is_conditional_on_the_existing_timeline_state() -> None:
    """10. It disappears the moment year columns exist, and it disappears
    because the APPLIED TIMELINE OWNER says so - not because anything maintains
    it."""
    structure = _structure()
    formula = structure.state_messages["inflation_formula"].replace(" ", "")
    assert formula.startswith('=IF(nmYearCount_Applied="",')
    # THE THREE OUTCOMES, AND THE THIRD IS SILENCE. A timeline is applied and
    # the span is non-empty: the cell is empty and the user sees nothing.
    assert formula.endswith('""))')
    assert structure.state_messages["inflation_empty_span"].replace(" ", "") in formula
    # AND THE ONLY NAME IT READS IS THE APPLIED-TIMELINE OWNER'S.
    names = set(re.findall(r"\bnm[A-Za-z0-9_]+", formula))
    assert names == {"nmYearCount_Applied", "nmInflFirstYear", "nmInflLastYear"}, names
    declared = {field.defined_name for field in structure.applied}
    declared |= {field.defined_name for field in structure.derived}
    assert names <= declared, sorted(names - declared)


def test_13_no_second_state_owner_or_macro_status_arrived_with_it() -> None:
    """11. THE ONE THING A PROMINENT MESSAGE MUST NOT BECOME.

    It is a formula in a cell, evaluated by Excel, reading names the structural
    owner already publishes. No VBA writes it, no VBA reads it, and no new
    PCCM_ procedure exists to maintain it.
    """
    structure = _structure()
    message = structure.state_messages["inflation_not_applied"]
    source = "\n".join(p.read_text(encoding="utf-8") for p in sorted(SRC.glob("*.bas")))
    source += (SRC / "ThisWorkbook.vba").read_text(encoding="utf-8")
    assert message not in source, "a module spells the state message"
    assert "inflation_not_applied" not in source
    for grid in structure.all_grids:
        assert f"B{grid.state_message_row}" not in source, grid.sheet
    # AND NO SECOND APPLY BUTTON APPEARED ON INFLATION.
    inflation = [b for b in structure.buttons if b.sheet == "Inflation"]
    assert inflation == [], inflation
    # NOR A NEW STATE VOCABULARY. The message is a sentence, not a state word:
    # nothing derives from it and nothing compares against it.
    assert message not in (SPEC / "sim_contract.yaml").read_text(encoding="utf-8")
    assert message not in (SPEC / "calc_contract.yaml").read_text(encoding="utf-8")


# ===========================================================================
# UX-004. ONE GEOMETRY FOR ALL FOUR
# ===========================================================================
def test_14_all_four_charts_have_identical_outer_dimensions() -> None:
    """1 and 2. THE CORRECTION ITSELF, and it is a single equality.

    UX-001 gave the tornado 20.0 cm against its neighbours' 16.0 to buy its
    driver names horizontal room. In real Excel that did not read as one chart
    with longer labels; it read as one chart being disproportionately large. The
    room did not have to come from outer width, and now it does not.

    ASSERTED AS A SET OF ONE, not as a tolerance. Two charts that differ by a
    millimetre are still visibly unequal on a rendered dashboard, and there is
    no reason for any of these four to differ at all.
    """
    widths = {chart["width_cm"] for chart in _projection()["charts"]}
    heights = {chart["height_cm"] for chart in _projection()["charts"]}
    assert len(widths) == 1, sorted(widths)
    assert len(heights) == 1, sorted(heights)
    # AND IT IS THE GEOMETRY THE WORKBOOK ACTUALLY CARRIES, not only the one the
    # manifest declares. The projection and the drawing are written by the same
    # build and have to agree.
    extents = {(round(chart.anchor.ext.cx / 360000, 2),
                round(chart.anchor.ext.cy / 360000, 2))
               for chart in _workbook()["Dashboard"]._charts}
    assert len(extents) == 1, sorted(extents)
    assert extents == {(widths.pop(), heights.pop())}, extents


def test_15_the_tornado_buys_its_label_room_inside_the_chart() -> None:
    """5. MEANINGFUL INTERNAL HORIZONTAL ROOM, stated in centimetres.

    Excel's automatic layout gives a category axis roughly a sixth of the chart
    and clips whatever does not fit - about 3 cm of driver name at this size.
    The manual allocation reserves the leftmost third, which is more absolute
    room than the 20 cm chart had, on a chart 2 cm narrower.

    AND ONLY THIS CHART DECLARES ONE. The other three plot years, bin edges and
    money against numeric axes, where the automatic layout is right and a fixed
    fraction would be an answer that stops tracking the data.
    """
    tornado = _by_key()["tornado"]
    area = tornado["plot_area"]
    assert area is not None, "the tornado has no declared plot area"
    assert set(area) == {"x", "y", "w", "h"}, sorted(area)
    # A THIRD OF THE CHART, AND AT LEAST AS MUCH AS EXCEL'S DEFAULT WOULD CLIP
    # AT. The fraction and the centimetres are both asserted: the fraction is
    # what the file carries, the centimetres are what a reader sees.
    assert area["x"] >= 0.30, area["x"]
    assert tornado["category_label_room_cm"] == round(area["x"] * tornado["width_cm"], 2)
    assert tornado["category_label_room_cm"] >= 5.0, tornado["category_label_room_cm"]
    # THE PLOT STILL GETS THE MAJORITY OF THE CHART. Label room taken so far it
    # left no plot would be a different defect, not a fix.
    assert area["w"] >= 0.55, area["w"]
    assert area["x"] + area["w"] <= 1.0
    assert area["y"] + area["h"] <= 1.0
    for key in CHART_KEYS:
        if key != "tornado":
            assert _by_key()[key]["plot_area"] is None, key
            assert _by_key()[key]["category_label_room_cm"] is None, key


def test_16_the_declared_plot_area_reaches_the_chart_part() -> None:
    """5. A layout openpyxl accepted and did not write would be an allocation
    that exists only in the projection. Exactly one chart part carries a manual
    layout, and it carries the declared fractions."""
    area = _by_key()["tornado"]["plot_area"]
    manual = [part for part in _chart_parts() if "<manualLayout>" in part]
    assert len(manual) == 1, f"{len(manual)} charts carry a manual plot area"
    part = manual[0]
    assert "Top Drivers by Rank Correlation" in part
    # EDGE-ANCHORED, or the fractions mean something else entirely: `edge` reads
    # them as a position in the chart, `factor` as an offset from where Excel
    # would have put it.
    assert '<xMode val="edge"' in part and '<yMode val="edge"' in part
    for name, value in area.items():
        assert f'<{name} val="{value}"' in part, (name, value)


def test_17_the_balanced_geometry_left_every_source_identity_alone() -> None:
    """6. UX-004 IS A SIZE CHANGE AND AN INTERNAL LAYOUT, NOTHING ELSE.

    Compared against the accepted UX batch rather than against the P8 tree, so
    this states what THIS correction did: no series, category, source block,
    state qualifier, title, axis format or no-data rule moved, and the bridge
    is untouched.
    """
    import subprocess
    accepted = yaml.safe_load(subprocess.run(
        ["git", "show", "0cf1b2d:pccm/spec/workbook.yaml"],
        cwd=PCCM_ROOT.parent, check=True,
        stdout=subprocess.PIPE).stdout.decode())["phase6_shell"]["charts"]
    was = {str(chart["key"]): chart for chart in accepted["charts"]}
    now = {str(chart["key"]): chart for chart in _charts()["charts"]}
    assert set(was) == set(now)
    # ONLY THESE THREE KEYS MAY DIFFER, and `plot_area` is the only new one.
    allowed = {"anchor", "width", "plot_area"}
    for key, before in was.items():
        after = now[key]
        assert set(after) - set(before) <= allowed, (key, set(after) - set(before))
        for field in set(before) - allowed:
            assert before[field] == after[field], (key, field)
    assert accepted["bridge"] == _charts()["bridge"], "the chart bridge moved"
    assert accepted["number_formats"] == _charts()["number_formats"]
    assert accepted["axis_presentation"] == _charts()["axis_presentation"]
    # AND THE HEIGHT DID NOT MOVE EITHER, which is why it is not in `allowed`.
    for key, before in was.items():
        assert before["height"] == now[key]["height"], key


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
