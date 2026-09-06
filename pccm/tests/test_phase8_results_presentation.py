#!/usr/bin/env python3
"""Phase 8, Step 1: the Results output surface, proved against its sources.

WHAT THIS SLICE IS. The two sections `Results` reserved from the beginning -
Annual Cash Flow and Reconciliation - stop being placeholders, on the sheet the
existing design already pointed at. No new worksheet, no new engine, and no VBA:
every cell is a lookup into the annual result Phase 7 published, or ordinary
display arithmetic over cells already on the sheet.

WHAT THESE CONTROLS EXIST TO CATCH. A presentation layer fails quietly. A
formula that averaged, interpolated or re-derived would look identical to one
that read, and a sheet that showed a HISTORICAL profile without saying so would
look exactly like a correct one. So the controls below check provenance, not
appearance: where each cell reads from, which block it reads, what it does when
there is nothing to show, and that no number the contracts own is typed here.

WHAT THEY CANNOT PROVE: there is no Excel here. Nothing below claims a formula
was evaluated, or that the sheet looks any particular way on screen.
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
from openpyxl import load_workbook  # noqa: E402

from pccm_builder import load_sim_contract  # noqa: E402
from pccm_builder.calc_loader import load_calc_contract  # noqa: E402
from pccm_builder.sim_emit import ANNUAL_PUBLISHED_MARKER  # noqa: E402
from pccm_builder.spec_loader import SpecError, load_spec  # noqa: E402
from pccm_builder.structure_loader import load_structure_contract  # noqa: E402

SPEC = PCCM_ROOT / "spec"
BUILD = PCCM_ROOT / "build"
SRC = PCCM_ROOT / "src" / "vba"

_CACHE: dict = {}


def _manifest() -> dict:
    if "manifest" not in _CACHE:
        _CACHE["manifest"] = json.loads(
            (BUILD / "stage_b_manifest.json").read_text(encoding="utf-8"))
    return _CACHE["manifest"]


def _sheet():
    if "sheet" not in _CACHE:
        workbook = load_workbook(BUILD / _manifest()["stage_a_filename"])
        _CACHE["workbook"] = workbook
        _CACHE["sheet"] = workbook["Results"]
    return _CACHE["sheet"]


def _workbook():
    _sheet()
    return _CACHE["workbook"]


def _shell() -> dict:
    if "shell" not in _CACHE:
        _CACHE["shell"] = yaml.safe_load(
            (SPEC / "workbook.yaml").read_text(encoding="utf-8"))["phase6_shell"]["results"]
    return _CACHE["shell"]


def _raw() -> dict:
    if "raw" not in _CACHE:
        _CACHE["raw"] = load_sim_contract(SPEC / "sim_contract.yaml").raw
    return _CACHE["raw"]


def _annual_contract() -> dict:
    return _raw()["sim_data"]["annual_records"]


def _window() -> int:
    return int(load_structure_contract(
        SPEC / "structure_contract.yaml").limits.max_generated_year_columns)


def _text_cells() -> list[str]:
    sheet = _sheet()
    found = []
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=10):
        for cell in row:
            if isinstance(cell.value, str) and not cell.value.startswith("="):
                found.append(cell.value)
    return found


def _formula(address: str) -> str:
    value = _sheet()[address].value
    assert isinstance(value, str) and value.startswith("="), (address, value)
    return value


# ===========================================================================
# A. PLACEMENT: THE SHEET THE DESIGN ALREADY POINTED AT
# ===========================================================================

def test_01_the_workbook_still_has_exactly_fourteen_sheets() -> None:
    """NO NEW WORKSHEET. `Annual Cash Flow` and `Reconciliation` are roadmap
    concepts, not sheets: the accepted `Results` sheet has declared "cash flow
    and the analytical reconciliation bridge" as its own purpose since Phase 1."""
    names = _workbook().sheetnames
    assert len(names) == 14, names
    for invented in ("Annual Cash Flow", "Reconciliation", "Cash Flow"):
        assert invented not in names, f"a {invented!r} worksheet was added"


def test_02_the_placement_authority_is_the_manifest_itself() -> None:
    spec = yaml.safe_load((SPEC / "workbook.yaml").read_text(encoding="utf-8"))
    results = [s for s in spec["sheets"] if s["name"] == "Results"]
    assert results, "the manifest declares no Results sheet"
    purpose = results[0]["purpose"].lower()
    assert "cash flow" in purpose and "reconciliation" in purpose, purpose
    titles = [b.get("title") for b in results[0]["blocks"]]
    assert "Annual Cash Flow" in titles and "Reconciliation" in titles, titles


def test_03_nothing_is_still_advertised_as_deferred() -> None:
    """The two sections exist now, so the sheet must stop saying they do not."""
    shell = _shell()
    assert "deferred" not in shell, "the deferred placeholders are still declared"
    joined = "\n".join(_text_cells()).lower()
    assert "deferred" not in joined, "the built sheet still calls something deferred"
    assert "Annual Cash Flow" in _text_cells()
    assert "Reconciliation" in _text_cells()


def test_04_no_later_phase_is_implied(self=None) -> None:
    """Phase 8 Step 1 is three sections. It must not advertise the dashboard,
    the charts or the Phase-9 model-check UI that come after it."""
    lowered = "\n".join(_text_cells()).lower()
    for later in ("dashboard", "s-curve", "histogram", "tornado", "chart",
                  "sensitivity", "spearman", "correlation", "warning", "repair"):
        assert later not in lowered, f"Results implies {later!r} exists here"


# ===========================================================================
# B. PROVENANCE: EVERY CELL READS SOMETHING SOMEBODY ELSE PUBLISHED
# ===========================================================================

def test_10_the_annual_table_reads_the_contracts_own_record_columns() -> None:
    annual = _annual_contract()
    shell = _shell()["annual"]
    sheet_name = _raw()["sim_data"]["sheet"]
    first = int(annual["first_record_row"])
    for offset in (0, 1, _window() - 1):
        row = int(shell["first_row"]) + offset
        for column in shell["columns"]:
            source = (annual["index_columns"] if column["source"] == "index"
                      else annual["selected_px_profile_columns"])
            formula = _formula(f"{column['column']}{row}")
            for bank in ("A", "B"):
                address = f"{sheet_name}!${source[bank][column['field']]}${first + offset}"
                assert address in formula, (column["key"], bank, address, formula)


def test_11_the_table_never_computes_anything() -> None:
    """A LOOKUP IS A LOOKUP. IF, OR and the bank switch are the only operators a
    record cell is allowed; an arithmetic operator here would be an annual
    engine on a worksheet, and no test of the VBA would ever see it."""
    shell = _shell()["annual"]
    banned = re.compile(r"AVERAGE|PERCENTILE|QUARTILE|STDEV|VAR|RAND|NORM|"
                        r"FORECAST|TREND|SLOPE|CORREL|RANK|LARGE|SMALL")
    for offset in range(0, _window(), 37):
        row = int(shell["first_row"]) + offset
        for column in shell["columns"]:
            formula = _formula(f"{column['column']}{row}")
            assert not banned.search(formula.upper()), (column["key"], formula)
            assert "*" not in formula and "/" not in formula, (column["key"], formula)
            assert "+" not in formula, (column["key"], formula)


def test_12_no_sim_data_address_is_written_in_the_manifest() -> None:
    """THE P7-4 DEFECT, CLOSED BY CONSTRUCTION. The sensitivity availability
    line spells its addresses as literals; when the persisted block moved from
    J/S to CC/CL the formula did not, and a successful run reported "Not
    produced for this run" forever. These blocks carry no address at all."""
    text = yaml.safe_dump(
        {"annual": _shell()["annual"], "reconciliation": _shell()["reconciliation"]},
        sort_keys=True)
    sheet_name = _raw()["sim_data"]["sheet"]
    assert sheet_name not in text, f"the Phase-8 blocks name {sheet_name} directly"
    assert not re.search(r"\$[A-Z]{1,3}\$\d+", text), (
        "the Phase-8 blocks carry an absolute cell address")


def test_13_no_state_word_is_typed_into_the_manifest() -> None:
    handoff = _annual_contract()["handoff"]
    text = yaml.safe_dump(
        {"annual": _shell()["annual"], "reconciliation": _shell()["reconciliation"]},
        sort_keys=True)
    for state in set(handoff["distribution_states"]) | set(handoff["profile_states"]):
        assert f'"{state}"' not in text and f"'{state}'" not in text, (
            f"the manifest spells the state word {state!r}; it is the contract's")
    assert ANNUAL_PUBLISHED_MARKER not in text


def test_14_the_publication_marker_has_one_owner() -> None:
    """The sheet decides whether an annual answer exists at all by comparing the
    stamp against this marker. A second copy of the string would let the sheet
    report NOT PRODUCED after a publication that had just succeeded."""
    emitted = (BUILD / "vba" / "modSimContract.bas").read_text(encoding="utf-8")
    assert (f'Public Const SIM_ANNUAL_PUBLISHED As String = '
            f'"{ANNUAL_PUBLISHED_MARKER}"') in emitted
    assert f'="{ANNUAL_PUBLISHED_MARKER}"' in _formula(_state_address("distribution_state"))


def _state_address(key: str) -> str:
    shell = _shell()
    return f"{shell['nominal_column']}{shell['annual'][f'{key}_row']}"


# ===========================================================================
# C. STATE DISPLAY: THE FOUR HONEST ANSWERS
# ===========================================================================

def test_20_the_distribution_state_uses_the_contracts_own_vocabulary() -> None:
    handoff = _annual_contract()["handoff"]
    formula = _formula(_state_address("distribution_state"))
    for state in handoff["distribution_states"]:
        assert f'"{state}"' in formula, (state, formula)
    assert f'"{handoff["inconsistent_stamp_state"]}"' not in formula, (
        "OTHER Px is a PROFILE state; the distributions never carry it")


def test_21_current_requires_the_stamp_to_be_this_runs_own_identity() -> None:
    """ALL FIVE FIELDS. A run id says which attempt; the fingerprint says which
    request and the digest says which answer. Any one of them alone would let a
    different run's annual result be presented as this one's."""
    raw = _raw()
    identity = {f["key"]: int(f["row"]) for f in raw["sim_data"]["run_identity"]["fields"]}
    stamp = {f["key"]: int(f["row"])
             for f in _annual_contract()["stamp"]["fields"]}
    formula = _formula(_state_address("distribution_state"))
    stamp_columns = _annual_contract()["stamp"]["bank_value_columns"]
    run_columns = raw["sim_data"]["run_identity"]["bank_value_columns"]
    sheet_name = raw["sim_data"]["sheet"]
    pairs = (("run_id", "run_id"), ("effective_seed", "effective_seed"),
             ("request_fingerprint", "request_fingerprint"),
             ("result_digest", "result_digest"), ("iterations", "iterations_run"))
    for stamp_key, run_key in pairs:
        for bank in ("A", "B"):
            assert f"{sheet_name}!${stamp_columns[bank]}${stamp[stamp_key]}" in formula, stamp_key
            assert f"{sheet_name}!${run_columns[bank]}${identity[run_key]}" in formula, run_key
    # AND THE SIMULATION ITSELF MUST STILL BE CURRENT.
    status_row = identity["simulation_status"]
    value_column = raw["sim_data"]["run_identity"]["value_column"]
    assert f"{sheet_name}!${value_column}${status_row}" in formula
    assert f'="{raw["sim_state"]["states"][0]}"' in formula


def test_22_the_profile_inherits_the_distribution_verdict_first() -> None:
    """A profile cannot be current for a run whose distributions are not, and it
    is never relabelled: what changes when the selector moves is only that
    nobody is asking for the Px it was computed at."""
    handoff = _annual_contract()["handoff"]
    shell = _shell()
    formula = _formula(_state_address("profile_state"))
    distribution_cell = f"${shell['nominal_column']}${shell['annual']['distribution_state_row']}"
    assert formula.startswith(f'=IF({distribution_cell}<>"{handoff["distribution_states"][1]}"'
                              f',{distribution_cell},'), formula
    assert f'"{handoff["inconsistent_stamp_state"]}"' in formula
    assert "inpSelectedConfidenceLevel" in formula, (
        "the selector arm must compare the STAMPED Px against the selected one")


def test_23_nothing_is_fabricated_when_nothing_was_produced() -> None:
    """NOT PRODUCED SHOWS NOTHING. A zero here would be a cash flow the model
    never produced, and it would sum, reconcile and chart like a real one."""
    handoff = _annual_contract()["handoff"]
    not_produced = handoff["distribution_states"][0]
    shell = _shell()
    guard = f'${shell["nominal_column"]}${shell["annual"]["distribution_state_row"]}="{not_produced}"'
    for offset in (0, _window() - 1):
        row = int(shell["annual"]["first_row"]) + offset
        for column in shell["annual"]["columns"]:
            formula = _formula(f"{column['column']}{row}")
            assert guard in formula, (column["key"], formula)
            assert ',"",' in formula, (column["key"], formula)
    for key in ("profile_px", "year_count"):
        assert guard in _formula(_state_address(key))


def test_24_historical_data_stays_visible_and_stays_labelled() -> None:
    """The contract keeps a superseded answer readable; the sheet must show it
    and must not let it read as the current one."""
    handoff = _annual_contract()["handoff"]
    historical = handoff["distribution_states"][-1]
    shell = _shell()
    row_formula = _formula(f"{shell['annual']['columns'][0]['column']}"
                           f"{shell['annual']['first_row']}")
    assert f'"{historical}"' not in row_formula, (
        "the rows blank on NOT PRODUCED only; a historical answer stays visible")
    state = _formula(_state_address("distribution_state"))
    assert f'"{historical}"' in state


# ===========================================================================
# D. THE ANNUAL WINDOW IS DYNAMIC
# ===========================================================================

def test_30_the_window_is_the_structural_maximum_not_a_tested_duration() -> None:
    shell = _shell()["annual"]
    window = _window()
    first = int(shell["first_row"])
    for column in shell["columns"]:
        assert _sheet()[f"{column['column']}{first + window - 1}"].value is not None
        beyond = _sheet()[f"{column['column']}{first + window}"].value
        assert beyond is None or not str(beyond).startswith("="), (
            f"the annual window reaches past the structural maximum of {window}")
    assert window >= 200, window


def test_31_every_row_is_bounded_by_the_stamped_year_count() -> None:
    """THE COUNT SAYS WHERE THE ANSWER STOPS - never the last non-blank row. A
    four-year run published over a twenty-year one leaves nothing behind."""
    shell = _shell()
    annual = shell["annual"]
    year_count_cell = f"${shell['nominal_column']}${annual['year_count_row']}"
    for offset in (0, 3, 19, _window() - 1):
        row = int(annual["first_row"]) + offset
        formula = _formula(f"{annual['columns'][0]['column']}{row}")
        assert f"{offset + 1}>{year_count_cell}" in formula, (offset, formula)


def test_32_no_duration_this_project_has_run_is_written_anywhere() -> None:
    """No hard-coded 4-year or 20-year display logic."""
    shell = yaml.safe_dump({"annual": _shell()["annual"],
                            "reconciliation": _shell()["reconciliation"]}, sort_keys=True)
    for duration in (" 4\n", " 20\n", " 200\n"):
        assert duration not in shell, f"the manifest hard-codes a duration:{duration!r}"
    source = (PCCM_ROOT / "builder" / "pccm_builder" / "workbook_builder.py").read_text(
        encoding="utf-8")
    annual_source = source[source.index("PHASE 8, STEP 1 - THE ANNUAL CASH FLOW"):]
    for literal in ("== 4", "== 20", "range(4)", "range(20)", "range(200)"):
        assert literal not in annual_source, f"the renderer hard-codes {literal!r}"


def test_33_the_year_count_comes_from_the_stamp() -> None:
    stamp = {f["key"]: int(f["row"]) for f in _annual_contract()["stamp"]["fields"]}
    columns = _annual_contract()["stamp"]["bank_value_columns"]
    sheet_name = _raw()["sim_data"]["sheet"]
    formula = _formula(_state_address("year_count"))
    for bank in ("A", "B"):
        assert f"{sheet_name}!${columns[bank]}${stamp['year_count']}" in formula


# ===========================================================================
# E. RECONCILIATION
# ===========================================================================

def _reconciliation_rows() -> dict[str, int]:
    reconciliation = _shell()["reconciliation"]
    return {entry["key"]: int(reconciliation["first_row"]) + index
            for index, entry in enumerate(reconciliation["rows"])}


def test_40_the_total_is_the_total_and_never_the_contingency() -> None:
    """THE W5 SEMANTIC, MADE PERMANENT. The summary block's quantile rungs are
    the percentile of the iteration TOTALS; the contingency block's rungs are
    `selected_px_total - deterministic_base_estimate_a`, a smaller number by
    exactly the base. Reading the second as the first failed four
    reconciliations by exactly that amount, and nothing on a sheet would show
    it: both are money, both move with the selector."""
    shell = _shell()
    rows = _reconciliation_rows()
    selected = shell["selected"]
    for column in (shell["nominal_column"], shell["pv_column"]):
        formula = _formula(f"{column}{rows['total']}")
        assert formula == f'=${column}${selected["quantile_row"]}', formula
        assert str(selected["contingency_row"]) not in formula.replace(
            str(selected["quantile_row"]), ""), (
            "the reconciliation reads the contingency row")


def test_41_the_selected_px_block_still_separates_the_two_ladders() -> None:
    """And the block it reads from is still the right one on _SimData."""
    raw = _raw()
    summary = raw["sim_data"]["summary_statistics"]["bank_value_columns"]
    contingency = raw["sim_data"]["contingency_ladder"]["bank_value_columns"]
    selected = _shell()["selected"]
    sheet_name = raw["sim_data"]["sheet"]
    for measure, key in (("nominal", "quantile_nominal"), ("pv", "quantile_pv")):
        formula = selected[key]
        for bank in ("A", "B"):
            assert f"{sheet_name}!${summary[bank][measure]}$" in formula, (measure, bank)
            assert f"{sheet_name}!${contingency[bank][measure]}$" not in formula
    for measure, key in (("nominal", "contingency_nominal"), ("pv", "contingency_pv")):
        formula = selected[key]
        for bank in ("A", "B"):
            assert f"{sheet_name}!${contingency[bank][measure]}$" in formula, (measure, bank)
            assert f"{sheet_name}!${summary[bank][measure]}$" not in formula


def test_42_the_sum_is_the_annual_profile_and_nothing_else() -> None:
    shell = _shell()
    rows = _reconciliation_rows()
    annual = shell["annual"]
    profile = {c["field"]: c["column"] for c in annual["columns"] if c["source"] == "profile"}
    first = int(annual["first_row"])
    last = first + _window() - 1
    for measure, column in (("nominal", shell["nominal_column"]), ("pv", shell["pv_column"])):
        formula = _formula(f"{column}{rows['profile_sum']}")
        assert f"SUM(${profile[measure]}${first}:${profile[measure]}${last})" in formula, formula


def test_43_the_allowance_is_the_projects_own_identity_rule() -> None:
    """I3c/I4c with ERRATUM C1: the conditioning scale names the MAGNITUDE OF
    THE ARITHMETIC PERFORMED - the annual terms summed plus the aggregate they
    are compared against - never the magnitude of the net result."""
    tolerances = load_calc_contract(SPEC / "calc_contract.yaml").tolerances
    shell = _shell()
    rows = _reconciliation_rows()
    for column in (shell["nominal_column"], shell["pv_column"]):
        formula = _formula(f"{column}{rows['allowance']}")
        for value in (tolerances.identity_absolute_floor,
                      tolerances.identity_relative_coefficient,
                      tolerances.conditioning_scale_floor):
            rendered = f"{value:.20f}".rstrip("0") if value < 1 else "1.0"
            assert rendered.rstrip(".") in formula or f"{value:g}" in formula, (
                value, formula)
        assert "MAX(" in formula and "ABS(" in formula
        assert "SUMIF(" in formula, "the conditioning scale must sum the terms' magnitudes"


def test_44_the_verdict_is_taken_on_the_stored_values() -> None:
    """The display may round; the decision may not. And a mismatch may not be
    hidden - the status says NOT RECONCILED, it does not blank."""
    shell = _shell()
    rows = _reconciliation_rows()
    verdicts = shell["reconciliation"]["verdicts"]
    for column in (shell["nominal_column"], shell["pv_column"]):
        formula = _formula(f"{column}{rows['status']}")
        assert f'ABS(${column}${rows["difference"]})<=${column}${rows["allowance"]}' in formula
        assert f'"{verdicts["reconciled"]}"' in formula
        assert f'"{verdicts["mismatch"]}"' in formula
        # AND SO MUST EVERY CELL THE VERDICT READS. Rounding the difference in
        # its own cell would round the comparison just as effectively as
        # rounding it inside the status - the status would simply be comparing a
        # number that had already lost the mismatch.
        for key in ("status", "difference", "allowance", "profile_sum", "total"):
            depended = _formula(f"{column}{rows[key]}")
            for rounding in ("ROUND(", "ROUNDUP(", "ROUNDDOWN(", "TRUNC(",
                             "TEXT(", "FIXED(", "INT("):
                assert rounding not in depended, (
                    f"the verdict depends on {key}, which rounds with {rounding}")
        assert formula is not None


def test_45_a_reconciliation_of_a_non_current_profile_says_so() -> None:
    shell = _shell()
    rows = _reconciliation_rows()
    handoff = _annual_contract()["handoff"]
    state_cell = f"${shell['nominal_column']}${shell['annual']['profile_state_row']}"
    for column in (shell["nominal_column"], shell["pv_column"]):
        formula = _formula(f"{column}{rows['status']}")
        assert f'IF({state_cell}="{handoff["profile_states"][1]}",""' in formula, formula
        assert shell["reconciliation"]["verdicts"]["qualifier_suffix"] in formula


def test_46_no_profile_is_scaled_to_make_the_identity_hold() -> None:
    shell = _shell()
    rows = _reconciliation_rows()
    for key in ("total", "profile_sum", "difference"):
        for column in (shell["nominal_column"], shell["pv_column"]):
            formula = _formula(f"{column}{rows[key]}")
            assert "*" not in formula and "/" not in formula, (key, formula)


# ===========================================================================
# F. PRESENTATION AND SAFETY RAILS
# ===========================================================================

def test_50_the_money_and_year_columns_are_formatted() -> None:
    shell = _shell()
    formats = shell["number_formats"]
    annual = shell["annual"]
    first = int(annual["first_row"])
    for column in annual["columns"]:
        assert _sheet()[f"{column['column']}{first}"].number_format == formats[column["format"]]
    for metric in shell["summary"]["metrics"]:
        for column in (shell["nominal_column"], shell["pv_column"]):
            assert _sheet()[f"{column}{metric['row']}"].number_format == formats["money"]
    rows = _reconciliation_rows()
    for entry in shell["reconciliation"]["rows"]:
        for column in (shell["nominal_column"], shell["pv_column"]):
            assert _sheet()[f"{column}{rows[entry['key']]}"].number_format == (
                formats[entry["format"]])


def test_50b_the_difference_can_show_a_difference_that_matters() -> None:
    """A FORMAT CAN HIDE A FAILURE AS EFFECTIVELY AS A ROUNDED COMPARISON. The
    verdict is taken unrounded, but a delta of 4.5e-13 displayed as `0` next to
    the word NOT RECONCILED would leave a reader with no way to see how far out
    it is - and a delta of 0.4 displayed as `0` would look like agreement. The
    two diagnostic rows must therefore be able to render a magnitude below the
    identity floor itself."""
    tolerances = load_calc_contract(SPEC / "calc_contract.yaml").tolerances
    floor = float(tolerances.identity_absolute_floor)
    formats = _shell()["number_formats"]
    for entry in _shell()["reconciliation"]["rows"]:
        if entry["key"] not in ("difference", "allowance"):
            continue
        pattern = formats[entry["format"]]
        if "E+" in pattern.upper():
            continue
        decimals = len(pattern.split(".")[1]) if "." in pattern else 0
        assert 10 ** -decimals <= floor, (
            f"{entry['key']} is formatted {pattern!r}, which cannot show a value "
            f"below the identity floor {floor!r}; a real mismatch would display as zero")


def test_51_no_machine_coordinate_or_field_name_is_shown_to_the_reader() -> None:
    shown = [t for t in _text_cells() if t not in ("Results",)]
    joined = "\n".join(shown)
    for machine in ("_SimData", "_Calc", "AB$", "$AZ", "bank_value_columns",
                    "project_index", "selected_px_label", "request_fingerprint"):
        assert machine not in joined, f"the sheet shows the machine name {machine!r}"


def test_52_the_headers_name_the_measures_in_words() -> None:
    annual = _shell()["annual"]
    headers = [column["header"] for column in annual["columns"]]
    assert headers == ["Project Year", "Calendar Year",
                       "Selected Px - Nominal (SAR)", "Selected Px - PV (SAR)"], headers
    header_row = int(annual["header_row"])
    for column in annual["columns"]:
        assert _sheet()[f"{column['column']}{header_row}"].value == column["header"]


def test_53_no_vba_was_added_for_this_slice() -> None:
    """PHASE 8 STEP 1 IS FORMULAS. The workbook already recalculates when
    `_SimData` changes, so a refresh engine would be a second presentation
    mechanism with nothing to do."""
    modules = sorted(p.name for p in SRC.glob("*.bas"))
    assert "modResults.bas" not in modules and "modPresentation.bas" not in modules
    for module in SRC.glob("*.bas"):
        text = module.read_text(encoding="utf-8")
        assert "PCCM_RefreshResults" not in text
        assert "Annual Cash Flow" not in text


def test_54_the_layout_refuses_to_overlap_the_accepted_phase_6_rows() -> None:
    """The Phase-6 coordinates are read by an accepted Windows harness. A
    Phase-8 section one row too high would overwrite one and every check of it
    would still pass, because both would read whatever landed there last."""
    spec_text = (SPEC / "workbook.yaml").read_text(encoding="utf-8")
    broken = spec_text.replace("      heading_row: 51", "      heading_row: 40", 1)
    path = PCCM_ROOT / "build" / "_phase8_layout_probe.yaml"
    path.write_text(broken, encoding="utf-8")
    try:
        with pytest.raises(SpecError, match="accepted Phase-6 layout"):
            load_spec(path)
    finally:
        path.unlink()


def test_55_the_window_may_not_grow_over_the_reconciliation() -> None:
    """The annual table is sized by a contract the Results layout does not own.
    If the structural year maximum grew past the space reserved for it, the
    later write would simply win and both sections would still 'render'."""
    from pccm_builder import workbook_builder

    source = (PCCM_ROOT / "builder" / "pccm_builder" / "workbook_builder.py").read_text(
        encoding="utf-8")
    assert "the structural year maximum grew past" in source
    shell = _shell()
    last = int(shell["annual"]["first_row"]) + _window() - 1
    assert last < int(shell["reconciliation"]["heading_row"]), (last, shell)
    assert hasattr(workbook_builder, "_annual_row_window")
