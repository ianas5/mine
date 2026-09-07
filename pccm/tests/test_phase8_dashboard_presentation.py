#!/usr/bin/env python3
"""Phase 8, Step 2: the Dashboard executive summary, proved against its sources.

WHAT THIS SLICE IS. The Dashboard stops being a placeholder and becomes a
one-screen executive summary. It adds no engine, no state rule and no VBA: every
analytical cell on it is `=IF(Results!$D$nn="","",Results!$D$nn)` and nothing
else. The workbook stays at fourteen sheets and no chart is drawn.

WHAT THESE CONTROLS EXIST TO CATCH. A second presentation layer fails in a
specific way: it re-derives something the layer below it already owns, the two
answers agree for a while, and then they do not. P8-1 spent a round on exactly
that - a Results sheet that rebuilt half the annual state in formulas and drifted
from the accessor that owned it. The Dashboard is one layer further out, so the
controls below check PROVENANCE first and appearance second: which cell each
value comes from, that Results is the only place it can come from, that the two
ladders and the two Px questions stay two, and that a blank stays blank.

WHAT THEY CANNOT PROVE: there is no Excel here. Nothing below claims a formula
was evaluated, that the sheet looks any particular way, or that a live state word
reached it. That is a Windows question and this step did not run Windows.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

import pytest  # noqa: E402
import yaml  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

from pccm_builder.phase8_dashboard import (  # noqa: E402
    build_phase8_dashboard_inspection,
    validate_phase8_dashboard_inspection,
)
from pccm_builder.spec_loader import SpecError, load_spec  # noqa: E402

SPEC = PCCM_ROOT / "spec"
BUILD = PCCM_ROOT / "build"
SRC = PCCM_ROOT / "src" / "vba"
MANIFEST = SPEC / "workbook.yaml"

# THE COMMIT P8-1 WAS ACCEPTED AT. Its Results geometry is the surface a Windows
# run has already been produced against; this step may extend the manifest but
# may not move a byte of that block.
P81_ACCEPTANCE = "35bd6ce"

_CACHE: dict = {}


def _shell() -> dict:
    if "shell" not in _CACHE:
        _CACHE["shell"] = yaml.safe_load(MANIFEST.read_text(encoding="utf-8"))["phase6_shell"]
    return _CACHE["shell"]


def _dashboard_block() -> dict:
    return _shell()["dashboard"]


def _results_block() -> dict:
    return _shell()["results"]


def _projection() -> dict:
    if "projection" not in _CACHE:
        _CACHE["projection"] = json.loads(
            (BUILD / "phase8_dashboard_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["projection"]


def _results_projection() -> dict:
    if "results_projection" not in _CACHE:
        _CACHE["results_projection"] = json.loads(
            (BUILD / "phase8_results_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["results_projection"]


def _workbook():
    if "workbook" not in _CACHE:
        manifest = json.loads((BUILD / "stage_b_manifest.json").read_text(encoding="utf-8"))
        _CACHE["workbook"] = load_workbook(BUILD / manifest["stage_a_filename"])
    return _CACHE["workbook"]


def _sheet():
    return _workbook()["Dashboard"]


def _formula(address: str) -> str:
    value = _sheet()[address].value
    assert isinstance(value, str) and value.startswith("="), (
        f"Dashboard!{address} holds {value!r}, not a formula")
    return value


def _entries() -> list[dict]:
    return [entry for section in _projection()["sections"] for entry in section["rows"]]


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=REPO_ROOT, check=True,
                          stdout=subprocess.PIPE, text=True).stdout


# ===========================================================================
# A. THE SHEET, AND WHAT IT IS NOT
# ===========================================================================
def test_01_the_workbook_still_has_exactly_fourteen_sheets() -> None:
    """The Dashboard was already there. P8-2 fills it in; it does not add a
    surface, and an executive summary that needed its own sheet would be a
    second output surface rather than a summary of the first."""
    names = _workbook().sheetnames
    assert len(names) == 14, names
    assert "Dashboard" in names
    assert names[0] == "Dashboard", (
        "the Dashboard is no longer the first sheet a reader opens")


def test_02_no_chart_exists_anywhere_in_the_workbook() -> None:
    """P8-2 DRAWS NOTHING. The S-curve, the histogram and the tornado belong to
    a later step; a chart object here would be that step started early and
    unreviewed."""
    for worksheet in _workbook().worksheets:
        charts = list(getattr(worksheet, "_charts", []))
        assert not charts, f"{worksheet.title} carries {len(charts)} chart(s)"
        images = list(getattr(worksheet, "_images", []))
        assert not images, f"{worksheet.title} carries {len(images)} image(s)"


def test_03_the_reserved_chart_region_is_declared_and_empty() -> None:
    """RESERVED MEANS RESERVED. A heading and a note say what the space is for;
    every cell inside it is empty, so the next step draws onto a blank region
    rather than over something."""
    region = _projection()["chart_region"]
    sheet = _sheet()
    heading = sheet[f"B{region['heading_row']}"].value
    assert isinstance(heading, str) and heading, "the chart region has no heading"
    assert isinstance(sheet[f"B{region['note_row']}"].value, str)
    for row in range(int(region["first_row"]), int(region["last_row"]) + 1):
        for column in "ABCDEFGH":
            value = sheet[f"{column}{row}"].value
            assert value is None, (
                f"Dashboard!{column}{row} is inside the reserved chart region but "
                f"holds {value!r}")


def test_04_no_dashboard_state_owner_was_added_to_the_vba() -> None:
    """FORMULA-ONLY. No module gained a Dashboard procedure, and the Phase-8 VBA
    surface is still the four Results adapters and nothing else."""
    for module in sorted(SRC.glob("*.bas")):
        text = module.read_text(encoding="utf-8")
        code = "\n".join(line for line in text.splitlines()
                         if not line.lstrip().startswith("'"))
        assert "Dashboard" not in code, (
            f"{module.name} names the Dashboard in executable code")
    adapter = (SRC / "modResultsState.bas").read_text(encoding="utf-8")
    assert len(re.findall(r"^Public Function (\w+)", adapter, re.M)) == 4


# ===========================================================================
# B. PROVENANCE - RESULTS IS THE ONLY SOURCE
# ===========================================================================
def test_10_every_analytical_cell_is_a_mirror_of_a_results_cell() -> None:
    """THE CENTRAL CLAIM, taken cell by cell over the whole sheet rather than
    over a sample. Each projected entry names the Results cell it copies; the
    workbook must hold exactly that reference and nothing else."""
    template = _projection()["mirror_formula"]
    for entry in _entries():
        for measure, cells in entry["cells"].items():
            address = cells["dashboard"]
            reference = "$".join(["", cells["results"].split("!")[0] + "!"])
            sheet, cell = cells["results"].split("!")
            column, row = re.match(r"([A-Z]+)(\d+)", cell).groups()
            expected = template.format(ref=f"{sheet}!${column}${row}")
            assert _formula(address) == expected, (
                f"Dashboard!{address} is {_formula(address)!r}, not the mirror of "
                f"{cells['results']}")
            assert reference  # the reference was parseable


def test_11_no_dashboard_formula_reaches_sim_data_or_calc() -> None:
    """NOT ONE MACHINE ADDRESS. `_SimData` is raw iteration output and `_Calc`
    is derived factors; an executive value taken from either would be bypassing
    the surface that gives it meaning, and would go stale the way P7-4's
    hand-written Sensitivity address did."""
    sheet = _sheet()
    for row in sheet.iter_rows():
        for cell in row:
            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                continue
            for forbidden in ("_SimData", "_Calc", "tbl", "inp"):
                assert forbidden not in cell.value, (
                    f"Dashboard!{cell.coordinate} reaches {forbidden}: {cell.value!r}")
            assert cell.value.count("Results!") == 2, (
                f"Dashboard!{cell.coordinate} does not read Results twice: "
                f"{cell.value!r}")


def test_12_the_dashboard_computes_nothing_at_all() -> None:
    """NO ARITHMETIC, NO COMPARISON, NO STATE WORD, NO VERDICT. A dashboard that
    subtracted a base from a total, summed a profile or compared a difference to
    an allowance would be a second implementation of an identity Results already
    owns - and the two would agree right up until they did not."""
    sheet = _sheet()
    banned = ("SUM", "SUMIF", "INDEX", "MATCH", "COUNTIF", "AVERAGE", "ABS",
              "MAX", "MIN", "ROUND", "VLOOKUP", "XLOOKUP", "IFERROR", "IFS",
              "AND(", "OR(", "TEXT(", "CONCAT", "&")
    for row in sheet.iter_rows():
        for cell in row:
            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                continue
            body = cell.value
            for token in banned:
                assert token not in body, (
                    f"Dashboard!{cell.coordinate} uses {token}: {body!r}")
            for operator in ("+", "-", "*", "/", "<", ">"):
                assert operator not in body, (
                    f"Dashboard!{cell.coordinate} performs arithmetic or a "
                    f"comparison: {body!r}")


def test_13_no_state_word_or_verdict_is_typed_on_the_dashboard() -> None:
    """A HARD-CODED "CURRENT" IS A LIE WAITING TO HAPPEN. Every state word and
    every reconciliation verdict arrives through a mirror or not at all - not
    from the manifest, not from a formula, not from a label."""
    words = ["CURRENT", "HISTORICAL", "NOT PRODUCED", "OTHER Px", "STALE",
             "INVALID", "Reconciled", "NOT RECONCILED"]
    sheet = _sheet()
    for row in sheet.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            if cell.value.startswith("="):
                for word in words:
                    assert word not in cell.value, (
                        f"Dashboard!{cell.coordinate} spells the state word {word!r}")
    manifest_block = json.dumps(_dashboard_block())
    for word in ("CURRENT", "HISTORICAL", "NOT PRODUCED", "OTHER Px",
                 "Reconciled", "NOT RECONCILED"):
        assert word not in manifest_block, (
            f"the Dashboard manifest block types the state word {word!r}")


def test_14_the_dashboard_owns_no_geometry_of_its_own() -> None:
    """ONE AUTHORITY. The Dashboard names a Results BLOCK and KEY; the row comes
    from the Results block. A Results row literal in the Dashboard manifest
    would be a second declaration, and P7-4 is what a second declaration costs
    when the first one moves."""
    block = _dashboard_block()
    results_rows = (
        {int(f["row"]) for f in _results_block()["run_stamp"]["fields"]}
        | {int(m["row"]) for m in _results_block()["summary"]["metrics"]}
        | {int(_results_block()["selected"][k]) for k in
           ("confidence_level_row", "quantile_row", "contingency_row")})
    for section in block["sections"]:
        for entry in section["rows"]:
            assert set(entry) <= {"key", "source_block", "source_key", "format",
                                  "emphasis"}, entry
            assert "row" not in entry, (
                f"{entry['key']} carries a row of its own; the Results block owns it")
    # AND NO RESULTS ROW NUMBER APPEARS ANYWHERE IN THE DASHBOARD BLOCK.
    declared = {int(section["row"]) for section in block["sections"]}
    for section in block["sections"]:
        declared |= {int(section["first_row"]) + offset
                     for offset in range(len(section["rows"]))}
    stray = sorted(results_rows & declared - set(range(1, 60)))
    assert not stray, f"the Dashboard block reuses Results row numbers {stray}"


def test_15_the_projection_and_the_renderer_resolve_the_same_cells() -> None:
    """TWO READERS OF ONE MANIFEST, checked against each other. The projection
    is what a test or a later runner reads; the renderer is what the workbook
    got. If they ever disagreed, every mapping control above would be checking
    the projection against itself."""
    for entry in _entries():
        for cells in entry["cells"].values():
            sheet, cell = cells["results"].split("!")
            column, row = re.match(r"([A-Z]+)(\d+)", cell).groups()
            assert f"{sheet}!${column}${row}" in _formula(cells["dashboard"])


# ===========================================================================
# C. THE DISTINCTIONS THAT MUST SURVIVE
# ===========================================================================
def test_20_the_total_and_the_contingency_stay_two_different_rows() -> None:
    """W5 READ ONE AS THE OTHER and failed four reconciliations by exactly the
    deterministic base. Nothing on this sheet recomputes the difference, so a
    dashboard that mirrored one row twice would show the same money under both
    names with nothing to contradict it."""
    mirrored = {(e["source_block"], e["source_key"]): e for e in _entries()}
    total = mirrored[("selected", "total")]
    contingency = mirrored[("selected", "contingency")]
    assert total["source_row"] != contingency["source_row"]
    assert total["row"] != contingency["row"]
    assert total["label"] != contingency["label"]
    # AND THEY ARE THE ROWS RESULTS CALLS BY THOSE NAMES.
    selected = _results_projection()["selected"]
    assert total["source_row"] == selected["total_row"]
    assert contingency["source_row"] == selected["contingency_row"]


def test_21_the_selected_level_and_the_published_profile_px_stay_apart() -> None:
    """TWO DIFFERENT QUESTIONS. One is the level the reader is asking for now;
    the other is the level the stored profile was actually blended at. P8-1's
    accepted Part B is the case where they differ - selector P50, profile still
    P80 - and a dashboard that read either from the other's row would make that
    state unobservable here."""
    mirrored = {(e["source_block"], e["source_key"]): e for e in _entries()}
    selected = mirrored[("selected", "confidence_level")]
    profile_px = mirrored[("state", "profile_px")]
    assert selected["source_row"] != profile_px["source_row"]
    assert selected["row"] != profile_px["row"]
    results = _results_projection()
    assert selected["source_row"] == results["selected"]["confidence_level_row"]
    assert profile_px["source_row"] == results["state"]["profile_px"]["row"]


def test_22_all_four_live_state_lines_are_mirrored_and_stay_distinct() -> None:
    """THE FOUR ANSWERS ARE FOUR ANSWERS. The distribution state and the profile
    state are separately derived by the owner - a moved selector retires the
    profile and must not retire the ladders - so collapsing them here would
    destroy the distinction one row above where it was made."""
    mirrored = {(e["source_block"], e["source_key"]): e for e in _entries()}
    state = _results_projection()["state"]
    rows = set()
    for key in ("distribution_state", "profile_state", "profile_px", "year_count"):
        entry = mirrored[("state", key)]
        assert entry["source_row"] == state[key]["row"], key
        rows.add(entry["row"])
    assert len(rows) == 4, "two state lines share a Dashboard row"


def test_23_the_persisted_status_is_shown_as_persisted_not_as_the_live_answer() -> None:
    """THE FIRST DEFECT P8-1 DISCLOSED, kept closed one layer out. Results row
    23 is the LAST EVALUATED simulation status - a persisted pair that an
    ordinary model change does not refresh - and Results labels it so. The
    Dashboard carries that label unchanged, shows it BELOW the four live state
    lines, and says in its own note that it is not a live answer."""
    mirrored = {(e["source_block"], e["source_key"]): e for e in _entries()}
    persisted = mirrored[("run_stamp", "simulation_status")]
    assert "last evaluated" in persisted["label"].lower(), persisted["label"]
    for key in ("distribution_state", "profile_state"):
        assert mirrored[("state", key)]["row"] < persisted["row"], (
            "the persisted status is shown above a live state line")
    section = next(s for s in _projection()["sections"] if s["key"] == "status")
    note = _sheet()[f"B{section['note_row']}"].value
    assert isinstance(note, str)
    assert "last evaluated" in note.lower() and "live" in note.lower(), note


def test_24_the_annual_table_itself_is_not_reproduced() -> None:
    """A SUMMARY, NOT A SECOND COPY. The 200-row window stays on Results; what
    an executive needs is whether a profile exists, whose Px it is, whether it
    is the current answer and whether it reconciles - and the state lines above
    already carry the first three."""
    annual = _results_projection()["annual"]
    mirrored_rows = {e["source_row"] for e in _entries()}
    window = set(range(int(annual["first_row"]),
                       int(annual["first_row"]) + int(annual["row_window"])))
    assert not (mirrored_rows & window), (
        "the Dashboard mirrors rows from the annual record window")
    assert len(_entries()) < 30, (
        f"{len(_entries())} mirrored rows is a report, not an executive summary")


# ===========================================================================
# D. DISPLAY BEHAVIOUR - BLANKS, HISTORY AND HONESTY
# ===========================================================================
def test_30_every_mirror_guards_the_blank_before_returning_it() -> None:
    """A FABRICATED ZERO IS THE WHOLE HAZARD. Excel reads an empty reference
    back as 0, so a bare `=Results!$D$47` prints a confident 0 beside the words
    NOT PRODUCED. The guard is what makes an absent result look absent."""
    template = _projection()["mirror_formula"]
    assert template.startswith('=IF({ref}=""'), template
    assert template.count("{ref}") == 2
    for entry in _entries():
        for cells in entry["cells"].values():
            body = _formula(cells["dashboard"])
            assert body.startswith("=IF("), body
            assert '="",""' in body.replace(" ", ""), (
                f"Dashboard!{cells['dashboard']} does not return a blank for a blank "
                f"source: {body!r}")


def test_31_no_dashboard_cell_carries_a_literal_zero_or_a_placeholder_value() -> None:
    """NOTHING IS PRE-FILLED. A zero, a dash or an "n/a" typed into a value cell
    would be indistinguishable from a computed one the moment somebody glanced
    at the sheet."""
    sheet = _sheet()
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is None or cell.column_letter == "B":
                continue
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue
            assert not isinstance(cell.value, (int, float)), (
                f"Dashboard!{cell.coordinate} holds the literal {cell.value!r}")
            assert str(cell.value).strip().lower() not in {
                "0", "-", "n/a", "na", "tbd", "0.0"}, cell.coordinate


def test_32_the_historical_and_other_px_qualifiers_arrive_intact() -> None:
    """A HISTORICAL ANSWER MAY BE SHOWN; IT MAY NOT BE SHOWN AS CURRENT. Results
    appends its own qualifier to the reconciliation verdict - "- for the OTHER
    Px profile, not the current model answer" - and the Dashboard mirrors the
    whole cell, so the qualifier cannot be trimmed off in transit."""
    mirrored = {(e["source_block"], e["source_key"]): e for e in _entries()}
    status = mirrored[("reconciliation", "status")]
    results = _results_projection()["reconciliation"]
    assert status["source_row"] == results["rows"]["status"]
    for measure in ("nominal", "pv"):
        body = _formula(status["cells"][measure]["dashboard"])
        # THE WHOLE CELL, NOT A SLICE OF IT. LEFT/MID/SUBSTITUTE over a verdict
        # is how a qualifier disappears without anybody deciding it should.
        for surgery in ("LEFT", "MID", "RIGHT", "SUBSTITUTE", "REPLACE", "TRIM"):
            assert surgery not in body, f"the verdict is edited in transit: {body!r}"
    # AND THE STATE LINES THEMSELVES ARE MIRRORED WHOLE, for the same reason.
    for key in ("distribution_state", "profile_state"):
        body = _formula(mirrored[("state", key)]["cells"]["nominal"]["dashboard"])
        assert "LEFT" not in body and "SUBSTITUTE" not in body


def test_33_no_error_suppression_anywhere_on_the_sheet() -> None:
    """A #VALUE! ON THIS SHEET IS INFORMATION. P8-1's whole first Windows run
    turned on being able to see one; IFERROR here would have hidden the defect
    that round existed to find."""
    sheet = _sheet()
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                for mask in ("IFERROR", "IFNA", "ISERROR", "ISNA", "ISERR"):
                    assert mask not in cell.value, (
                        f"Dashboard!{cell.coordinate} suppresses an error: "
                        f"{cell.value!r}")


def test_34_number_formats_match_the_kind_of_thing_each_row_holds() -> None:
    """MONEY READS AS MONEY AND A VERDICT READS AS TEXT. A money format on a
    status cell would render a word as a number format's idea of one; a General
    format on a total would print fifteen significant figures at an executive."""
    formats = _projection()["number_formats"]
    sheet = _sheet()
    expected_kind = {
        "money": formats["money"], "year": formats["year"],
        "delta": formats["delta"], "text": formats["text"],
    }
    for entry in _entries():
        for cells in entry["cells"].values():
            actual = sheet[cells["dashboard"]].number_format
            assert actual == expected_kind[entry["format"]], (
                f"Dashboard!{cells['dashboard']} is {actual!r}, expected "
                f"{expected_kind[entry['format']]!r} for a {entry['format']} row")
    # AND THE KINDS THEMSELVES ARE THE RESULTS ONES, so the same number is not
    # rounded differently on two sheets.
    assert formats == _results_projection()["number_formats"]


def test_35_the_difference_is_not_rounded_into_looking_like_zero() -> None:
    """THE DELTA FORMAT IS SCIENTIFIC ON PURPOSE. A `#,##0` difference of
    4.5e-13 displays as 0 and a reader concludes the identity held exactly. It
    did not; it held within an allowance, and those are different statements."""
    mirrored = {(e["source_block"], e["source_key"]): e for e in _entries()}
    difference = mirrored[("reconciliation", "difference")]
    assert difference["format"] == "delta"
    fmt = _projection()["number_formats"]["delta"]
    assert "E+" in fmt, fmt
    for cells in difference["cells"].values():
        assert _sheet()[cells["dashboard"]].number_format == fmt


# ===========================================================================
# E. THE ACCEPTED SURFACES BELOW ARE UNTOUCHED
# ===========================================================================
def test_40_the_accepted_p8_1_results_geometry_did_not_move() -> None:
    """A WINDOWS RUN WAS PRODUCED AGAINST THAT BLOCK. P8-2 may extend the
    manifest; it may not move a byte of the surface P8-1 was accepted on. Read
    from the acceptance commit rather than asserted."""
    accepted = yaml.safe_load(
        _git("show", f"{P81_ACCEPTANCE}:pccm/spec/workbook.yaml"))
    assert accepted["phase6_shell"]["results"] == _results_block(), (
        "the accepted P8-1 Results block changed in this step")


def test_41_no_production_vba_changed_in_this_step() -> None:
    """FORMULA-ONLY MEANS FORMULA-ONLY."""
    changed = _git("diff", "--name-only", f"{P81_ACCEPTANCE}..HEAD", "--", "pccm/src")
    assert changed.strip() == "", f"production VBA changed in P8-2: {changed}"


def test_42_the_dashboard_is_declared_everywhere_a_surface_must_be() -> None:
    """THE OMISSION THAT COST THREE ROUNDS. `modResultsState` landed without
    being named in two Phase-5 inventories and they were red on the branch for
    four commits. A new PRESENTATION surface has its own inventories - the
    permitted-formula set and the projection - and this control refuses one that
    exists in the workbook but is declared in neither."""
    from pccm_builder.verify import _phase6_formula_cells
    from pccm_builder.structure_loader import load_structure_contract

    spec = load_spec(MANIFEST)
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    permitted = _phase6_formula_cells(spec, structure)
    assert "Dashboard" in permitted, (
        "the Dashboard writes formulas but is not in the permitted-formula set; "
        "structural verification would fail or, worse, be widened to pass")
    projected = {cells["dashboard"]
                 for entry in _entries() for cells in entry["cells"].values()}
    assert projected == permitted["Dashboard"], (
        "the permitted-formula set and the projection disagree about which cells "
        "the Dashboard writes")
    # AND THE WORKBOOK AGREES WITH BOTH.
    actual = {cell.coordinate for row in _sheet().iter_rows() for cell in row
              if isinstance(cell.value, str) and cell.value.startswith("=")}
    assert actual == projected, (
        f"the sheet writes {sorted(actual - projected)} that nothing declares, and "
        f"declares {sorted(projected - actual)} it does not write")


# ===========================================================================
# F. MUTATIONS - each is a mistake somebody would actually make
# ===========================================================================
def _mutated_spec(mutate) -> Path:
    raw = yaml.safe_load(MANIFEST.read_text(encoding="utf-8"))
    mutate(raw["phase6_shell"]["dashboard"], raw["phase6_shell"]["results"])
    path = BUILD / "_phase8_dashboard_probe.yaml"
    path.write_text(yaml.safe_dump(raw, sort_keys=False), encoding="utf-8")
    return path


@pytest.mark.parametrize("name,mutate,expected", [
    # THE BLANK GUARD REMOVED - the fabricated-zero defect, introduced.
    ("the mirror stops guarding the blank",
     lambda dash, res: dash.__setitem__("mirror_formula", "={ref}"),
     "guard the reference against blank"),
    # THE TOTAL POINTED AT THE CONTINGENCY - the W5 defect, one sheet out.
    ("the total mirrors the contingency row",
     lambda dash, res: dash["sections"][1]["rows"].__setitem__(
         1, {"key": "selected_total", "source_block": "selected",
             "source_key": "contingency", "format": "money", "emphasis": True}),
     None),
    # THE SELECTED LEVEL READ FROM THE PROFILE'S Px - Part B made unobservable.
    ("the selected level mirrors the published profile Px",
     lambda dash, res: dash["sections"][0]["rows"].__setitem__(
         4, {"key": "selected_confidence_level", "source_block": "state",
             "source_key": "profile_px", "format": "text"}),
     None),
    # A SECTION GROWN INTO THE RESERVED REGION.
    ("a section reaches into the chart region",
     lambda dash, res: dash["sections"][4].__setitem__("first_row", 52),
     "chart region reserved from row"),
    # A SOURCE KEY RESULTS DOES NOT PUBLISH - the P7-4 stale-address shape.
    ("a row mirrors a key Results does not publish",
     lambda dash, res: dash["sections"][1]["rows"][0].__setitem__(
         "source_key", "expected_risk_exposure"),
     "which the Results summary block does not publish"),
    # TWO SECTIONS OVERLAID - the silent one.
    ("two sections write the same row",
     lambda dash, res: dash["sections"][2].__setitem__("row", 16),
     "both write row"),
    # A FORMAT NOTHING DECLARES.
    ("a row names an undeclared number format",
     lambda dash, res: dash["sections"][1]["rows"][0].__setitem__(
         "format", "currency"),
     "no format is declared for"),
    # THE DASHBOARD MADE ITS OWN SOURCE.
    ("the dashboard mirrors itself",
     lambda dash, res: dash.__setitem__("source_sheet", "Dashboard"),
     "mirrors 'Results' and nothing else"),
])
def test_50_the_manifest_refuses_each_ownership_mistake(
        name: str, mutate, expected: str | None) -> None:
    """EACH MUTATION APPLIED TO A COPY OF THE MANIFEST, and refused by the
    loader or - for the two that are structurally legal but semantically wrong -
    by the projection's own validator. Nothing on disk changes."""
    path = _mutated_spec(mutate)
    try:
        if expected is not None:
            with pytest.raises(SpecError, match=re.escape(expected)):
                load_spec(path)
            return
        # STRUCTURALLY LEGAL, SEMANTICALLY WRONG. The loader cannot know that
        # mirroring the contingency under the total's name is a lie; the
        # projection validator can, because it compares the two rows.
        spec = load_spec(path)
        with pytest.raises(ValueError):
            validate_phase8_dashboard_inspection(
                build_phase8_dashboard_inspection(spec))
    finally:
        path.unlink(missing_ok=True)


def test_51_the_mutation_harness_is_not_vacuous() -> None:
    """THE UNMUTATED MANIFEST PASSES BOTH GATES, so the eight refusals above are
    refusals of the mutation and not of the fixture."""
    spec = load_spec(MANIFEST)
    inspection = build_phase8_dashboard_inspection(spec)
    validate_phase8_dashboard_inspection(inspection)
    assert inspection["sections"], "the projection is empty"
    assert inspection == _projection(), (
        "the committed projection is not what the manifest currently produces")
