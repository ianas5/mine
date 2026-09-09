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

# THE FOUR ANNUAL ADAPTERS P8-1 SETTLED. P8-3 adds a fifth, named where it is
# asserted rather than folded in here, so the two families stay distinguishable.
ADAPTERS = (
    "PCCM_ResultsAnnualDistributionState",
    "PCCM_ResultsAnnualProfileState",
    "PCCM_ResultsAnnualProfilePx",
    "PCCM_ResultsAnnualYearCount",
)

# THE MODULE'S WHOLE PUBLIC SURFACE, IN ORDER, as later phases have extended it.
# Named rather than counted, so an addition has to be declared here to pass and
# a rename or a reorder still fails.
LATER_ADAPTERS = ADAPTERS + (
    "PCCM_ResultsSimulationState",        # P8-3, the live simulation state
    "PCCM_ModelCheckCalculationState",    # P9-2, the live calculation state
    "PCCM_ModelCheckRefusalDetail",       # P9-2A, the live reason it is invalid
    "PCCM_ModelCheckRefusalSubject",      # P9-2B, which driver that reason is about
)

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


def _state_address(key: str) -> str:
    shell = _shell()
    return f"{shell['nominal_column']}{shell['annual'][f'{key}_row']}"


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
    """Phase 8 Step 1 is three sections. It must not advertise a phase that has
    not happened.

    NARROWED AT P8-3, AND DISCLOSED. The original list banned every word for
    every later step, which was right while every one of them was later. Two
    have now landed: P8-2 built the Dashboard and P8-3 put a chart bridge on
    THIS sheet, so those rows legitimately name what they are for - a block
    called 'Chart Data' that did not say it was for the charts would be worse,
    not better. What is still banned is what is still ahead: the Phase-9
    model-check UI and the Phase-10 hardening language. The P8-3 words are
    permitted only in the rows the chart projection declares, which is what the
    second half of this control checks."""
    banned_everywhere = ("warning", "repair", "remediation", "sign-off")
    lowered = "\n".join(_text_cells()).lower()
    for later in banned_everywhere:
        assert later not in lowered, f"Results implies {later!r} exists here"
    # AND THE CHART VOCABULARY IS CONFINED TO THE CHART BRIDGE. A P8-1 row that
    # started talking about charts would be the accepted surface drifting.
    charts = json.loads(
        (BUILD / "phase8_charts_inspection.json").read_text(encoding="utf-8"))
    bridge_top = int(charts["bridge"]["heading_row"])
    sheet = _sheet()
    for row in range(1, bridge_top):
        for column in "BDFHJ":
            value = sheet[f"{column}{row}"].value
            if not isinstance(value, str) or value.startswith("="):
                continue
            for later in ("dashboard", "s-curve", "histogram", "tornado", "chart",
                          "spearman"):
                assert later not in value.lower(), (
                    f"Results!{column}{row} names {later!r} above the chart bridge")


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


def test_14_the_state_cells_call_the_phase_7_accessors_through_an_adapter() -> None:
    """THE CORRECTION. Each state cell is a call and nothing else - no marker,
    no identity comparison, no selector arm, no state word. The adapter names
    are derived from the accessor names, so a renamed accessor cannot leave a
    stale wrapper on the sheet."""
    accessors = [str(e["name"]) for e in _annual_contract()["handoff"]["accessors"]]
    keys = ("distribution_state", "profile_state", "profile_px", "year_count")
    assert len(accessors) == len(keys), accessors
    for key, accessor in zip(keys, accessors):
        adapter = "PCCM_Results" + accessor[len("PCCM_"):]
        assert _formula(_state_address(key)) == f"={adapter}()", key


def test_15_the_adapter_owns_no_state_rule_at_all() -> None:
    """A WRAPPER, NOT A SECOND ENGINE. The module may call the accessors and do
    nothing else with a state: no state word, no marker, no stamp comparison."""
    module = (SRC / "modResultsState.bas").read_text(encoding="utf-8")
    code = "\n".join(line for line in module.splitlines()
                     if not line.lstrip().startswith("'"))
    handoff = _annual_contract()["handoff"]
    for state in set(handoff["distribution_states"]) | set(handoff["profile_states"]):
        assert f'"{state}"' not in code, f"the adapter spells the state word {state!r}"
    assert f'"{ANNUAL_PUBLISHED_MARKER}"' not in code
    for owned in ("StampText", "SharedText", "IsPublished", "StrComp", "Select Case",
                  "SIM_ANNUAL", "SIM_IDENTITY"):
        assert owned not in code, f"the adapter reaches into the store through {owned}"
    called = set(re.findall(r"modSimAnnualStore\.(PCCM_\w+)", code))
    assert called == {str(e["name"]) for e in handoff["accessors"]}, sorted(called)


def test_16_the_worksheet_holds_no_copy_of_the_state_decision_tree() -> None:
    """THE ROOT ISSUE, CHECKED WHERE IT WAS. Not one formula on the sheet may
    decide currentness: no publication-marker comparison, no stamp-against-run
    identity test, and no persisted simulation status read."""
    raw = _raw()
    sheet_name = raw["sim_data"]["sheet"]
    stamp = {f["key"]: int(f["row"]) for f in _annual_contract()["stamp"]["fields"]}
    stamp_columns = _annual_contract()["stamp"]["bank_value_columns"]
    identity = {f["key"]: int(f["row"]) for f in raw["sim_data"]["run_identity"]["fields"]}
    value_column = raw["sim_data"]["run_identity"]["value_column"]
    forbidden = {f"{sheet_name}!${value_column}${identity['simulation_status']}":
                 "the persisted simulation status"}
    for key in ("published", "run_id", "effective_seed", "request_fingerprint",
                "result_digest", "iterations"):
        for bank in ("A", "B"):
            forbidden[f"{sheet_name}!${stamp_columns[bank]}${stamp[key]}"] = (
                f"the annual stamp's {key}")
    # THE ACCEPTED RUN STAMP IS NOT A VERDICT. It reports the persisted status
    # under its own label - "Simulation Status (last evaluated)" - which is what
    # that block has always done and is exactly the caveat the reader needs. The
    # ban is on DECIDING with these addresses, so the run-stamp rows, whose rows
    # the manifest declares, are read out of the sweep rather than excused.
    stamp_rows = {int(f["row"]) for f in _shell()["run_stamp"]["fields"]}
    sheet = _sheet()
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=10):
        for cell in row:
            if not isinstance(cell.value, str) or not cell.value.startswith("="):
                continue
            if cell.row in stamp_rows:
                continue
            for address, what in forbidden.items():
                # ANCHORED AT THE ROW NUMBER. `$AB$8` is a prefix of `$AB$80`,
                # which is an ordinary record cell 72 rows into the annual
                # window - matching on the substring would convict the table of
                # reading the stamp it never touches.
                assert not re.search(re.escape(address) + r"(?!\d)", cell.value), (
                    f"{cell.coordinate} decides state from {what}; the accessor owns it")
    # And the run stamp reports it as LAST EVALUATED, never as the annual state.
    labels = [f["label"] for f in _shell()["run_stamp"]["fields"]
              if f["key"] == "simulation_status"]
    assert labels and "last evaluated" in labels[0].lower(), labels
    assert f'"{ANNUAL_PUBLISHED_MARKER}"' not in "\n".join(
        str(c.value) for r in sheet.iter_rows() for c in r if c.value is not None)


def test_17_the_renderer_can_no_longer_compose_a_verdict() -> None:
    source = (PCCM_ROOT / "builder" / "pccm_builder" / "workbook_builder.py").read_text(
        encoding="utf-8")
    block = source[source.index("def _annual_state_formulas("):
                   source.index("def _render_annual_section(")]
    for decision in ("stamp(", "run(", "AND(", "NOT(", "simulation_status",
                     "distribution_current", "historical", "other_px"):
        assert decision not in block, (
            f"the state formula builder still composes a verdict using {decision!r}")
    assert "state_procedures()" in block


# ===========================================================================
# C. STATE DISPLAY: THE FOUR HONEST ANSWERS, ASKED OF THEIR OWNER
# ===========================================================================

def test_20_the_owner_still_owns_every_arm_of_the_rule() -> None:
    """THE SECOND DISCLOSED DEFECT, CLOSED. The old formula implemented the
    selector arm of ProfileStateOf and not the stamp-consistency arm; now the
    whole function decides, and this control fails if the store stops carrying
    either arm - which would mean the sheet had quietly lost one again."""
    store = (SRC / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    profile = store[store.index("Private Function ProfileStateOf"):]
    profile = profile[:profile.index("Private Function StampBelongsTo")]
    # THE GUARDS, NOT THE NAMES. A parameter that is still declared and never
    # tested is exactly what the old worksheet formula was: an arm that exists
    # on paper and decides nothing.
    for guard in ("If Not stampConsistent Then",
                  "If Not selectorResolved Then",
                  "If StrComp(distribution, SIM_ANNUAL_STATE_CURRENT, "
                  "vbBinaryCompare) <> 0 Then"):
        assert guard in profile, f"the arm guarded by `{guard}` no longer decides anything"
    assert "SIM_ANNUAL_STATE_OTHER_PX" in profile
    assert "ProfileStateOf = distribution" in profile, (
        "the profile no longer inherits the distribution verdict")
    assert "If False Then" not in profile, "an arm was short-circuited"
    caller = store[store.index("Public Function PCCM_AnnualProfileState"):]
    caller = caller[:caller.index("Public Function PCCM_AnnualProfilePx")]
    assert "SimStatsSelectedProbability" in caller, (
        "the stamped label is no longer checked against its own probability")


def test_21_the_distribution_verdict_is_the_stores_and_reaches_the_live_status() -> None:
    """THE FIRST DISCLOSED DEFECT, STILL CLOSED. The banner used to read the
    PERSISTED status, so an ordinary model change left it saying CURRENT until
    some later operation re-evaluated it. The accessor derives the status.

    RESTATED, NOT WEAKENED, AFTER THE P8-1 READ-ONLY CORRECTION. What that
    correction changed is which ENTRY POINT the read path asks through - the one
    that does not also rewrite D28:D29 - and this control now names it. The
    claim is unchanged: the state cell is answered from a FRESH derivation, not
    from the stored word. Both entry points return DeriveSimStatus()."""
    store = (SRC / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    assert "SimAnnualStoreCurrentRunReadOnly(run, detail)" in store
    read_path = store[store.index("Public Function SimAnnualStoreCurrentRunReadOnly"):]
    read_path = read_path[:read_path.index("End Function")]
    assert "modSimReport.SimReportDerivedStatus()" in read_path, (
        "the store no longer asks for a freshly derived status")
    report = (SRC / "modSimReport.bas").read_text(encoding="utf-8")
    for entry, follower in (("Public Function SimReportDerivedStatus", "End Function"),
                            ("Public Function PCCM_SimulationStatus", "End Function")):
        body = report[report.index(entry):]
        body = body[:body.index(follower)]
        assert "DeriveSimStatus()" in body, f"{entry} no longer re-derives the status"
    # AND THE STORED WORD IS STILL NOT WHAT ANY OF THEM RETURNS.
    assert "SharedText(SIM_IDENTITY_ROW_SIMULATION_STATUS)" not in store, (
        "the store reads the persisted status word again")


def test_22_the_adapters_are_volatile_and_only_the_adapters_are() -> None:
    """VOLATILITY IS THE MECHANISM AND IT IS FENCED. A zero-argument function
    has no inputs Excel can watch, so without this it would answer once and keep
    answering. Nothing heavier is made volatile: not the calculation, not the
    simulation, not the annual run, not sensitivity."""
    adapter = (SRC / "modResultsState.bas").read_text(encoding="utf-8")
    functions = re.findall(r"^Public Function (\w+)", adapter, re.M)
    # FIVE SINCE P8-3 AND SIX SINCE P9-2, EACH ONE NAMED. The chart layer needed
    # a live SIMULATION state: the four annual ones read NOT PRODUCED whenever
    # the annual step has not run, which says nothing about a histogram whose
    # data is present. Model Check then needed a live CALCULATION state, for the
    # same two reasons one module along - the persisted `_Calc` C19 row is
    # last-evaluated, and the owner's own entry point persists and so may not be
    # called from a cell.
    #
    # THE CONTROL DID NOT WEAKEN WHEN IT GREW. It said "the count is asserted
    # with the names so a sixth cannot arrive unremarked", and a sixth arrived
    # and is remarked. A SEVENTH still fails, and so does a rename, a reorder or
    # a removal - which is more than a count on its own could ever say.
    assert tuple(functions) == LATER_ADAPTERS, functions
    assert adapter.count("Application.Volatile True") == len(functions), (
        "every adapter must be volatile, and there must be nothing else to make volatile")
    for module in sorted(SRC.glob("*.bas")):
        if module.name == "modResultsState.bas":
            continue
        text = module.read_text(encoding="utf-8")
        code = "\n".join(line for line in text.splitlines()
                         if not line.lstrip().startswith("'"))
        assert "Application.Volatile" not in code, (
            f"{module.name} was made volatile; only the presentation adapters may be")


def test_23_an_unavailable_accessor_fails_loud_rather_than_wrong() -> None:
    """A WRONG STATE WORD IS INDISTINGUISHABLE FROM A RIGHT ONE. Two of the four
    accessors reach a function that persists the derived status rows, and Excel
    does not let a function called from a cell change the workbook. If that ever
    raises rather than being ignored, the cell must show an error - never a
    plausible state."""
    adapter = (SRC / "modResultsState.bas").read_text(encoding="utf-8")
    functions = tuple(re.findall(r"^Public Function (\w+)", adapter, re.M))
    assert functions == LATER_ADAPTERS, functions
    count = len(functions)
    assert adapter.count("On Error GoTo Unavailable") == count
    assert adapter.count("CVErr(xlErrValue)") == count
    assert "Resume Next" not in adapter, (
        "swallowing the error would let the adapter return an empty state")


def test_24_no_endpoint_is_invoked_to_refresh_presentation_state() -> None:
    """The display asks what the state IS. It never runs anything to find out."""
    adapter = (SRC / "modResultsState.bas").read_text(encoding="utf-8")
    for endpoint in ("PCCM_Calculate", "PCCM_RunSimulation", "PCCM_RunAnnualStochastic",
                     "PCCM_RunSensitivity", "PCCM_AutomationBegin"):
        assert endpoint not in adapter, f"the adapter invokes {endpoint}"
    sheet = _sheet()
    formulas = "\n".join(str(c.value) for r in sheet.iter_rows() for c in r
                         if isinstance(c.value, str) and c.value.startswith("="))
    for endpoint in ("PCCM_Calculate(", "PCCM_RunSimulation(",
                     "PCCM_RunAnnualStochastic(", "PCCM_RunSensitivity("):
        assert endpoint not in formulas, f"a cell invokes {endpoint}"


def test_25_nothing_is_fabricated_when_nothing_was_produced() -> None:
    """NOT PRODUCED SHOWS NOTHING, and the verdict it blanks on is the
    accessor's. A zero here would be a cash flow the model never produced, and
    it would sum, reconcile and later chart like a real one."""
    handoff = _annual_contract()["handoff"]
    not_produced = handoff["distribution_states"][0]
    shell = _shell()
    guard = (f'${shell["nominal_column"]}${shell["annual"]["distribution_state_row"]}'
             f'="{not_produced}"')
    for offset in (0, _window() - 1):
        row = int(shell["annual"]["first_row"]) + offset
        for column in shell["annual"]["columns"]:
            formula = _formula(f"{column['column']}{row}")
            assert guard in formula, (column["key"], formula)
            assert ',"",' in formula, (column["key"], formula)


def test_26_historical_and_other_px_stay_visible_and_stay_labelled() -> None:
    """The contract keeps a superseded answer readable and never relabels a
    profile; the sheet blanks the window on NOT PRODUCED alone, and the state
    line beside it says which of the four it is."""
    handoff = _annual_contract()["handoff"]
    shell = _shell()
    row_formula = _formula(f"{shell['annual']['columns'][0]['column']}"
                           f"{shell['annual']['first_row']}")
    for still_visible in (handoff["distribution_states"][-1],
                          handoff["inconsistent_stamp_state"]):
        assert f'"{still_visible}"' not in row_formula, (
            f"the record window blanks on {still_visible}; it must stay readable")
    assert _sheet()[f"{shell['label_column']}"
                    f"{shell['annual']['profile_state_row']}"].value == (
        shell["annual"]["labels"]["profile_state"])
    # AND THE STAMPED Px IS THE STAMP'S, whatever the selector says now.
    assert _formula(_state_address("profile_px")).endswith("ProfilePx()")


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


def test_33_the_year_count_comes_from_the_handoff_authority() -> None:
    """AND THE STAMP IS STILL WHAT IT READS - in the accessor, where the rule
    that a count says where the answer stops has always lived."""
    accessors = [str(e["name"]) for e in _annual_contract()["handoff"]["accessors"]]
    assert _formula(_state_address("year_count")) == (
        f"=PCCM_Results{accessors[3][len('PCCM_'):]}()")
    store = (SRC / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    body = store[store.index(f"Public Function {accessors[3]}"):]
    body = body[:body.index("' =====")] if "' =====" in body else body[:1200]
    assert "SIM_ANNUAL_STAMP_ROW_YEAR_COUNT" in body, (
        "the year count accessor no longer reads the stamped count")


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


# ===========================================================================
# THE WORKSHEET PATH IS READ-ONLY - THE DEFECT THE FIRST COMPLETE RUN FOUND
# ===========================================================================
# WHAT HAPPENED, FROM A LIVE EXCEL. Part 0 evaluated all four state cells. From
# Part A onward - the moment a publication existed - the two STATE cells showed
# #VALUE! and the annual table and reconciliation cascaded off them.
#
# THE CHAIN, IN FOUR HOPS:
#     Results!D53   =PCCM_ResultsAnnualDistributionState()
#       -> modResultsState.PCCM_ResultsAnnualDistributionState
#       -> modSimAnnualStore.PCCM_AnnualDistributionState
#       -> modSimAnnualStore.SimAnnualStoreCurrentRun     (as it then was)
#       -> modSimReport.PCCM_SimulationStatus
#       -> modSimReport.WriteStatusBlock
#       -> SimSheet.Range("D28:D29").Value2 = block       <- PROHIBITED
#
# WHY PART 0 SURVIVED IT. Both accessors read the active publication bank first
# and return NOT PRODUCED when there is none. With no bank the guard exits
# before the precondition is ever asked for, so the write was never attempted.
# A publication makes the guard pass, and the very next thing the accessor does
# is ask a question that used to rewrite two cells.
#
# WHY THE OTHER TWO CELLS NEVER FAILED. PCCM_AnnualProfilePx and
# PCCM_AnnualYearCount read the stamp and stop. They never reach the
# precondition, so they never reached the write - which is exactly the pattern
# the live report showed, and it is the strongest single piece of evidence for
# the diagnosis.
#
# THE FIX IS AN OWNERSHIP SPLIT, NOT A SUPPRESSION. DeriveSimStatus was already
# pure; only its caller persisted. modSimReport now exposes the pure half under
# its own name, the store has a read-only precondition beside the command one,
# and the two accessors take the read-only one. No state rule moved, nothing was
# duplicated, and no error is being swallowed anywhere.


# EVERY WAY THIS PROJECT'S VBA CHANGES A WORKBOOK. Assignment through a Range or
# a Cells, a ListObject row operation, and the two clearing verbs. In-memory
# `Scripting.Dictionary.Add` is deliberately NOT here: it writes to a variable,
# not to the book, and treating it as a mutation would make the control cry wolf
# in modInflation and modStructuralCheck for no reason.
_WORKBOOK_WRITE = re.compile(
    r"""(?:\.Value2|\.Value|\.Formula\w*|\.NumberFormat|\.Text)\s*=(?!=)"""
    r"""|\.ClearContents\b|\.EntireRow\.Delete\b|ListRows\.Add\b|\.ListRows\("""
    r"""|Application\.(?:Calculation|EnableEvents|ScreenUpdating)\s*=""")


def _vba_procedures() -> dict[tuple[str, str], str]:
    """Every procedure in every production module, keyed (module, name), with
    comments and string literals removed so a WORD in prose is never a call."""
    if "vba_procs" not in _CACHE:
        from pccm_builder.vba_source import load_modules

        out: dict[tuple[str, str], str] = {}
        for module in load_modules([SRC]):
            lines = module.code.splitlines()
            index = 0
            while index < len(lines):
                head = re.match(
                    r"\s*(?:Public |Private |Friend )?(?:Static )?"
                    r"(?:Function|Sub|Property (?:Get|Let|Set))\s+([A-Za-z_]\w*)",
                    lines[index])
                if head:
                    end = index + 1
                    while end < len(lines) and not re.match(
                            r"\s*End (?:Function|Sub|Property)\b", lines[end]):
                        end += 1
                    out[(module.name, head.group(1))] = "\n".join(lines[index + 1:end])
                    index = end
                index += 1
        _CACHE["vba_procs"] = out
    return _CACHE["vba_procs"]  # type: ignore[return-value]


def _callees(module: str, body: str) -> set[tuple[str, str]]:
    """DELIBERATELY OVER-INCLUSIVE. A qualified `modX.Name` is a call; a bare
    identifier that names a procedure of the SAME module is treated as one even
    where it is a variable of that name. Over-inclusion can only ever make a
    read-only claim harder to satisfy, never easier, which is the direction a
    safety control has to err in."""
    procedures = _vba_procedures()
    found: set[tuple[str, str]] = set()
    for match in re.finditer(r"\b(mod\w+)\.([A-Za-z_]\w*)", body):
        found.add((match.group(1), match.group(2)))
    for match in re.finditer(r"\b([A-Za-z_]\w*)\b", body):
        if (module, match.group(1)) in procedures:
            found.add((module, match.group(1)))
    return found


def _reachable(root: tuple[str, str]) -> set[tuple[str, str]]:
    procedures = _vba_procedures()
    seen: set[tuple[str, str]] = set()
    stack = [root]
    while stack:
        current = stack.pop()
        if current in seen:
            continue
        seen.add(current)
        body = procedures.get(current)
        if body is None:
            continue
        for callee in _callees(current[0], body):
            if callee in procedures and callee not in seen:
                stack.append(callee)
    return seen


def _writes_reachable_from(root: tuple[str, str]) -> list[str]:
    procedures = _vba_procedures()
    out: list[str] = []
    for procedure in sorted(_reachable(root)):
        for line in procedures.get(procedure, "").splitlines():
            if _WORKBOOK_WRITE.search(line):
                out.append(f"{procedure[0]}.{procedure[1]}: {line.strip()[:80]}")
    return out


def test_60_the_call_graph_is_read_and_the_helper_is_not_vacuous() -> None:
    """THE INSTRUMENT, BEFORE THE MEASUREMENT. A reachability control that
    silently found no procedures would pass forever, so this proves the graph is
    populated, that it crosses modules, and that it does detect a write where one
    genuinely is."""
    procedures = _vba_procedures()
    assert len(procedures) > 400, len(procedures)
    assert ("modResultsState", "PCCM_ResultsAnnualDistributionState") in procedures
    assert ("modSimReport", "WriteStatusBlock") in procedures
    # IT CROSSES MODULES.
    reached = _reachable(("modResultsState", "PCCM_ResultsAnnualDistributionState"))
    assert ("modSimAnnualStore", "PCCM_AnnualDistributionState") in reached
    assert len({module for module, _ in reached}) > 3, reached
    # AND IT FINDS THE WRITE THAT IS STILL THERE, on the command path.
    writes = _writes_reachable_from(("modSimReport", "PCCM_SimulationStatus"))
    assert any("WriteStatusBlock" in entry for entry in writes), writes


@pytest.mark.parametrize("adapter", ADAPTERS)
def test_61_no_worksheet_adapter_can_reach_a_workbook_write(adapter: str) -> None:
    """THE CENTRAL CLAIM, TAKEN OVER THE WHOLE TRANSITIVE CLOSURE rather than
    over the adapter's own four lines. A cell may not change the book, so
    nothing a cell can call may either - however many hops away it is."""
    writes = _writes_reachable_from(("modResultsState", adapter))
    assert not writes, (
        f"{adapter} can reach a workbook write:\n  " + "\n  ".join(writes))


@pytest.mark.parametrize("adapter,persisting", [
    (a, p) for a in ADAPTERS
    for p in (("modSimReport", "PCCM_SimulationStatus"),
              ("modSimReport", "WriteStatusBlock"),
              ("modCalcReport", "PCCM_CalculationStatus"),
              ("modCalcReport", "WriteStatusBlock"))
])
def test_62_no_worksheet_adapter_reaches_a_persisting_entry_point(
        adapter: str, persisting: tuple[str, str]) -> None:
    """NAMED, NOT INFERRED. The two status endpoints and the two writers behind
    them are the specific procedures the live failure went through; each is
    asserted unreachable by name so a re-connection is caught even if some later
    refactor made the write itself unrecognisable to the pattern above."""
    assert persisting not in _reachable(("modResultsState", adapter)), (
        f"{adapter} reaches {persisting[0]}.{persisting[1]} again")


def test_63_the_command_path_still_persists_and_is_unchanged() -> None:
    """THE OTHER HALF OF THE SPLIT. An operation invoked from a button or an
    endpoint still derives AND persists: that is accepted Phase-7 behaviour and
    the correction must not have quietly removed it."""
    store = (SRC / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    run = (SRC / "modSimAnnualRun.bas").read_text(encoding="utf-8")
    assert "modSimAnnualStore.SimAnnualStoreCurrentRun(run, detail)" in run, (
        "the annual run no longer uses the command precondition")
    command = store[store.index("Public Function SimAnnualStoreCurrentRun"):]
    command = command[:command.index("End Function")]
    assert "modSimReport.PCCM_SimulationStatus()" in command, (
        "the command path no longer asks through the persisting entry point")
    assert ("modSimReport", "WriteStatusBlock") in _reachable(
        ("modSimAnnualRun", "RunAnnual")), (
        "the annual run no longer reaches the status persistence it always did")
    report = (SRC / "modSimReport.bas").read_text(encoding="utf-8")
    status = report[report.index("Public Function PCCM_SimulationStatus"):]
    status = status[:status.index("End Function")]
    assert "WriteStatusBlock status" in status, "PCCM_SimulationStatus stopped persisting"


def test_64_there_is_exactly_one_derivation_and_both_paths_return_it() -> None:
    """NO SECOND STATUS ENGINE. The split is about who WRITES, not about what
    the answer is: one private derivation, two public entry points over it, and
    the read-only one adds nothing of its own."""
    report = (SRC / "modSimReport.bas").read_text(encoding="utf-8")
    code = "\n".join(line for line in report.splitlines()
                     if not line.lstrip().startswith("'"))
    assert code.count("Private Function DeriveSimStatus") == 1, (
        "there is not exactly one simulation-status derivation")
    pure = code[code.index("Public Function SimReportDerivedStatus"):]
    pure = pure[:pure.index("End Function")]
    assert "DeriveSimStatus()" in pure
    for forbidden in ("WriteStatusBlock", "SIM_STATE_", "If ", "Range", ".Value"):
        assert forbidden not in pure, (
            f"the read-only status entry point does more than delegate: {forbidden}")
    # AND THE STORE SETTLES BOTH PATHS IN ONE PLACE.
    store = (SRC / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    store_code = "\n".join(line for line in store.splitlines()
                           if not line.lstrip().startswith("'"))
    assert store_code.count("Private Function CurrentRunFor") == 1
    assert store_code.count("CurrentRunFor(") == 3, (
        "the two preconditions no longer share one settlement")
    assert re.findall(r"SIM_STATE_\w+", store_code) == ["SIM_STATE_CURRENT"], (
        "the store carries a second simulation-state vocabulary")


def test_65_the_adapter_still_owns_no_state_rule_after_the_correction() -> None:
    """THE CORRECTION HAPPENED IN THE OWNER, NOT IN THE PRESENTATION. Nothing
    was copied down into modResultsState to route around the write."""
    adapter = (SRC / "modResultsState.bas").read_text(encoding="utf-8")
    code = "\n".join(line for line in adapter.splitlines()
                     if not line.lstrip().startswith("'"))
    # A RULE IS A DECISION, NOT A DELEGATION. The adapter may CALL the owner of
    # a semantic; what it may never do is decide one - no constant, no
    # comparison, no marker test, no state word of its own.
    #
    # `SimReportDerivedStatus` left this list at P8-3 and that is the whole
    # point of the correction: the fifth adapter calls it, exactly as the other
    # four call their accessors. `DeriveSimStatus` - the private derivation
    # itself - stays banned, because reaching past the owner into its internals
    # would be the duplication this control exists for.
    for rule in ("SIM_ANNUAL_STATE_", "SIM_STATE_", "CURRENT", "HISTORICAL",
                 "OTHER", "NOT PRODUCED", "STALE", "INVALID", "StrComp",
                 "DeriveSimStatus", "StampText", "SimAnnualStoreCurrentRun",
                 "WriteStatusBlock", "If ", "Select Case"):
        assert rule not in code, f"the adapter has acquired a state rule: {rule}"
    assert tuple(re.findall(r"^Public Function (\w+)", code, re.M)) == LATER_ADAPTERS
    # AND IT STILL CALLS THE ACCESSORS, one each, unchanged.
    for accessor in ("PCCM_AnnualDistributionState", "PCCM_AnnualProfileState",
                     "PCCM_AnnualProfilePx", "PCCM_AnnualYearCount"):
        assert code.count(f"modSimAnnualStore.{accessor}()") == 1, accessor
    # P8-3'S FIFTH DELEGATES TO THE ACCEPTED PURE OWNER AND TO NOTHING ELSE.
    # `SimReportDerivedStatus` is the read-only half P8-1 split out;
    # `PCCM_SimulationStatus` derives the same answer AND persists it, which a
    # worksheet cell may not do.
    assert code.count("modSimReport.SimReportDerivedStatus()") == 1
    assert "PCCM_SimulationStatus" not in code, (
        "the adapter reaches the writing status path")


@pytest.mark.parametrize("name,mutate,expect", [
    # RECONNECT THE UDF TO THE WRITING PATH - the defect itself, restored.
    ("the read path is pointed back at the persisting precondition",
     lambda store, report: (store.replace("SimAnnualStoreCurrentRunReadOnly(run, detail)",
                                          "SimAnnualStoreCurrentRun(run, detail)"), report),
     "write"),
    # MAKE THE READ-ONLY ENTRY POINT PERSIST TOO - the split in name only.
    ("the read-only entry point persists as well",
     lambda store, report: (store, report.replace(
         "    SimReportDerivedStatus = DeriveSimStatus()",
         "    SimReportDerivedStatus = DeriveSimStatus()\n    WriteStatusBlock "
         "SimReportDerivedStatus")),
     "write"),
    # DUPLICATE A STATE ARM INSIDE THE STORE'S READ PATH - a second authority.
    ("a second state vocabulary appears in the store",
     lambda store, report: (store.replace(
         "Public Function SimAnnualStoreCurrentRunReadOnly",
         "Public Function SimAnnualStoreExtraState()\n"
         "    SimAnnualStoreExtraState = SIM_STATE_STALE\n"
         "End Function\n\n"
         "Public Function SimAnnualStoreCurrentRunReadOnly"), report),
     "vocabulary"),
])
def test_66_the_mutations_that_would_undo_the_correction_are_refused(
        name: str, mutate, expect: str) -> None:
    """EACH MUTATION APPLIED TO A COPY OF THE TWO MODULES IN A TEMPORARY TREE,
    with the controls above re-run over it. Nothing on disk changes."""
    import shutil
    import tempfile

    from pccm_builder.vba_source import load_modules

    with tempfile.TemporaryDirectory() as raw:
        temp = Path(raw) / "vba"
        shutil.copytree(SRC, temp)
        store_path = temp / "modSimAnnualStore.bas"
        report_path = temp / "modSimReport.bas"
        store, report = mutate(store_path.read_text(encoding="utf-8"),
                               report_path.read_text(encoding="utf-8"))
        store_path.write_text(store, encoding="utf-8")
        report_path.write_text(report, encoding="utf-8")
        assert (store, report) != (SRC.joinpath("modSimAnnualStore.bas").read_text(
            encoding="utf-8"), SRC.joinpath("modSimReport.bas").read_text(
            encoding="utf-8")), f"the mutation '{name}' changed nothing"

        procedures: dict[tuple[str, str], str] = {}
        for module in load_modules([temp]):
            lines = module.code.splitlines()
            index = 0
            while index < len(lines):
                head = re.match(
                    r"\s*(?:Public |Private |Friend )?(?:Static )?"
                    r"(?:Function|Sub|Property (?:Get|Let|Set))\s+([A-Za-z_]\w*)",
                    lines[index])
                if head:
                    end = index + 1
                    while end < len(lines) and not re.match(
                            r"\s*End (?:Function|Sub|Property)\b", lines[end]):
                        end += 1
                    procedures[(module.name, head.group(1))] = "\n".join(
                        lines[index + 1:end])
                    index = end
                index += 1

        def callees(module: str, body: str) -> set[tuple[str, str]]:
            found: set[tuple[str, str]] = set()
            for match in re.finditer(r"\b(mod\w+)\.([A-Za-z_]\w*)", body):
                found.add((match.group(1), match.group(2)))
            for match in re.finditer(r"\b([A-Za-z_]\w*)\b", body):
                if (module, match.group(1)) in procedures:
                    found.add((module, match.group(1)))
            return found

        def reachable(root):
            seen, stack = set(), [root]
            while stack:
                current = stack.pop()
                if current in seen:
                    continue
                seen.add(current)
                for callee in callees(current[0], procedures.get(current, "")):
                    if callee in procedures and callee not in seen:
                        stack.append(callee)
            return seen

        if expect == "write":
            offending = [
                adapter for adapter in ADAPTERS
                if any(_WORKBOOK_WRITE.search(line)
                       for procedure in reachable(("modResultsState", adapter))
                       for line in procedures.get(procedure, "").splitlines())]
            assert offending, f"'{name}' left every adapter read-only"
        else:
            code = "\n".join(line for line in store.splitlines()
                             if not line.lstrip().startswith("'"))
            assert re.findall(r"SIM_STATE_\w+", code) != ["SIM_STATE_CURRENT"], (
                f"'{name}' did not introduce a second state vocabulary")


def test_67_the_four_required_state_semantics_are_still_the_owners() -> None:
    """THE SEMANTICS THE CORRECTION MUST NOT HAVE MOVED. Every arm the live
    parts depend on - NOT PRODUCED, CURRENT, OTHER Px, HISTORICAL - is still
    decided by the same two private rules in the same module, and the accessors
    still feed them the same three facts."""
    store = (SRC / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    for rule, arms in (
            ("Private Function DistributionStateOf",
             ("SIM_ANNUAL_STATE_NOT_PRODUCED", "SIM_ANNUAL_STATE_CURRENT",
              "SIM_ANNUAL_STATE_HISTORICAL")),
            ("Private Function ProfileStateOf",
             ("ProfileStateOf = distribution", "SIM_ANNUAL_STATE_OTHER_PX",
              "SIM_ANNUAL_STATE_CURRENT"))):
        body = store[store.index(rule):]
        body = body[:body.index("\nEnd Function")]
        for arm in arms:
            assert arm in body, f"{rule} lost the arm {arm!r}"
    # THE THREE FACTS, AND THE READ-ONLY PRECONDITION AMONG THEM.
    for accessor in ("PCCM_AnnualDistributionState", "PCCM_AnnualProfileState"):
        body = store[store.index(f"Public Function {accessor}"):]
        body = body[:body.index("\nEnd Function")]
        assert "SIM_ANNUAL_STAMP_ROW_PUBLISHED" in body
        assert "SimAnnualStoreCurrentRunReadOnly(run, detail)" in body, (
            f"{accessor} is not on the read-only precondition")
        assert "StampBelongsTo(bank, run)" in body
        assert "If Not IsBank(bank) Then" in body, (
            f"{accessor} lost the guard that makes Part 0 answer NOT PRODUCED")
    # AND THE TWO THAT NEVER FAILED STILL NEVER ASK THE PRECONDITION.
    for accessor in ("PCCM_AnnualProfilePx", "PCCM_AnnualYearCount"):
        body = store[store.index(f"Public Function {accessor}"):]
        body = body[:body.index("\nEnd Function")]
        assert "CurrentRun" not in body, f"{accessor} has acquired the precondition"


# EVERYTHING ELSE A WORKSHEET FUNCTION MAY NOT DO. A write is what broke this
# one, but it is not the only prohibition, and the next one would present
# identically: an error object in a cell and nothing to say which line raised.
# So the read path is swept for the whole class, not just the member that bit.
_UDF_HOSTILE = re.compile(
    r"\.Select\b|\.Activate\b|MsgBox|Application\.Run\b|\.Calculate\b"
    r"|Application\.(?:ScreenUpdating|EnableEvents|Calculation|DisplayAlerts)\s*="
    r"|ListRows\.Add|\.EntireRow|\.Delete\b|\.Copy\b|\.PasteSpecial|SendKeys"
    r"|\.SaveAs\b|ActiveWorkbook|ActiveSheet|Selection\b|DoEvents")


@pytest.mark.parametrize("adapter", ADAPTERS)
def test_68_the_read_path_does_nothing_else_a_cell_is_forbidden_to_do(
        adapter: str) -> None:
    """THE WHOLE CLASS, NOT THE ONE MEMBER. Selecting, activating, running an
    endpoint, forcing a calculation, changing an application setting, showing a
    dialog, or touching the active book instead of a named one - each is
    prohibited to a function called from a cell, and each would fail exactly the
    way the write did."""
    procedures = _vba_procedures()
    offenders = [
        f"{module}.{name}: {line.strip()[:80]}"
        for module, name in sorted(_reachable(("modResultsState", adapter)))
        for line in procedures.get((module, name), "").splitlines()
        if _UDF_HOSTILE.search(line)]
    assert not offenders, (
        f"{adapter} can reach an operation a worksheet function may not "
        f"perform:\n  " + "\n  ".join(offenders))


def test_69_the_hostile_sweep_is_not_vacuous() -> None:
    """THE SWEEP ABOVE PROVES A NEGATIVE, so it has to be shown capable of a
    positive. Each pattern is matched against a line that genuinely contains
    it, and the graph it runs over is shown to be populated."""
    for line in ("    ws.Select", "    Application.Run \"PCCM_Calculate\"",
                 "    Application.ScreenUpdating = False", "    MsgBox detail",
                 "    ActiveSheet.Range(\"A1\").Value2 = 1", "    DoEvents"):
        assert _UDF_HOSTILE.search(line), line
    for line in ("    value = Sh(SIM_DATA_SHEET).Range(\"D30\").Value2",
                 "    If Len(status) = 0 Then Exit Function"):
        assert not _UDF_HOSTILE.search(line), line
    assert len(_reachable(("modResultsState", ADAPTERS[1]))) > 100
