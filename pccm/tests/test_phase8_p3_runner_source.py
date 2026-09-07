#!/usr/bin/env python3
"""P8-3: the minimal Windows runner for the analytical chart layer.

WHAT IT IS FOR. P8-3 drew four charts and proved them statically. Whether the
chart OBJECTS Excel actually loads are the projected ones, wired to the
projected bridge ranges, plotting what those cells hold - and drawing NOTHING
where the bridge says there is no point - is a runtime question nothing on Linux
can answer.

WHAT THESE CONTROLS EXIST TO CATCH. A chart runner fails in three specific ways
and each is silent. It can recompute what it is checking - a second histogram, a
reconstructed cumulative series, a re-ranked tornado - and then agree with itself
about a chart that has drifted. It can read the chart before freezing the cells
it will be compared against, so an accessor manufactures the agreement. And it
can accept a zero where the bridge said #N/A, which is the fabricated year this
whole layer was built to prevent.

WHAT THIS FILE CANNOT PROVE: there is no PowerShell and no Excel here, so
nothing below claims the runner RAN.
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

import test_phase7_acceptance_harness_source as accepted  # noqa: E402

WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
RUNNER = WINDOWS / "phase8_p3_chart_surface.ps1"
P81 = WINDOWS / "phase8_p1_results_surface.ps1"
P82 = WINDOWS / "phase8_p2_dashboard_surface.ps1"
TIMING = WINDOWS / "phase7_timing_scenarios.ps1"
LIFECYCLE = WINDOWS / "com_lifecycle.ps1"
PHASE5 = WINDOWS / "phase5_gate_b_scenarios.ps1"
PHASE6 = WINDOWS / "phase6_gate_b_scenarios.ps1"
BUILD = PCCM_ROOT / "build"

DOT_SOURCED = (LIFECYCLE, PHASE5, PHASE6)

COPIED_HELPERS = (
    "Write-RowObject", "Get-NamedValue", "Set-NamedValue", "Get-TableColumnNames",
    "Set-TableCell", "Get-TableBody", "Get-TableRowCount", "Add-BlankTableRow",
    "Remove-TableRow", "Get-IdColumnValues",
)

# THE PARTS REUSED FROM THE ACCEPTED P8-2 RUNNER, renamed and otherwise
# untouched. Reusing a lifecycle that has run on Windows means reusing it.
SHARED_WITH_P82 = (
    "Write-P83Line", "Add-P83Check", "Invoke-P83Release", "Get-P83SourceRevision",
    "Invoke-P83Endpoint", "Set-P83NamedText", "Get-P83Register",
    "Get-P83RegisterColumnIndex", "Get-P83RegisterRowIndex", "Get-P83Cell",
    "Format-P83Cell", "Test-P83Blank", "Test-P83SameCellValue", "Invoke-P83Recalculate",
)

BANNERS = (
    "PART 0 - NOTHING HAS RUN",
    "PART A1 - CALCULATE, SIMULATE, ANNUAL - AND NO SENSITIVITY",
    "PART A2 - THE SENSITIVITY ENDPOINT, INVOKED BY THIS RUNNER",
    "PART B - THE SELECTOR MOVES, RECALCULATION ONLY",
    "PART C - THE ANNUAL ENDPOINT ALONE",
    "PART D - ITERATIONS CHANGE, RECALCULATION ONLY",
    "PART E - AN INVALID MODEL",
)

_CACHE: dict = {}


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    return accepted._ps_code(RUNNER)


def _split_functions(code: str) -> dict[str, str]:
    """THE FUNCTION BODIES OF ANY COMMENT-FREE COPY OF THE RUNNER, so a rule in
    section F can be re-run over a mutated copy the same way it runs over disk."""
    out: dict[str, str] = {}
    for match in re.finditer(r"^function\s+([\w-]+)\s*\{", code, re.M):
        start = match.end() - 1
        depth = 0
        for index in range(start, len(code)):
            if code[index] == "{":
                depth += 1
            elif code[index] == "}":
                depth -= 1
                if depth == 0:
                    out[match.group(1)] = code[start:index]
                    break
    return out


def _functions() -> dict[str, str]:
    if "fns" not in _CACHE:
        _CACHE["fns"] = _split_functions(_code())
    return _CACHE["fns"]


def _function(name: str) -> str:
    body = _functions().get(name)
    assert body is not None, f"{name} is not defined in the P8-3 runner"
    return body


def _part(name: str) -> str:
    code = _code()
    starts = []
    for banner in BANNERS:
        marker = f"Write-P83Line '{banner}'"
        assert marker in code, f"the runner has no banner for {banner!r}"
        starts.append(code.index(marker))
    assert starts == sorted(starts), starts
    index = {"part 0": 0, "part A1": 1, "part A2": 2, "part B": 3,
             "part C": 4, "part D": 5, "part E": 6}[name]
    end = starts[index + 1] if index + 1 < len(starts) else len(code)
    return code[starts[index]:end]


def _projection() -> dict:
    if "charts" not in _CACHE:
        _CACHE["charts"] = json.loads(
            (BUILD / "phase8_charts_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["charts"]


# ===========================================================================
# A. SCOPE AND REUSE
# ===========================================================================
def test_01_there_is_one_dedicated_p8_3_runner() -> None:
    assert RUNNER.is_file()
    assert _text().startswith("<#"), "the runner has no synopsis"
    assert "P8-3" in _text()
    # AND IT IS NOT THE EARLIER RUNNERS EDITED. Their accepted evidence stands
    # at their own commits and neither may be repurposed.
    others = sorted(path.name for path in WINDOWS.glob("phase8_*.ps1"))
    assert others == ["phase8_p1_results_surface.ps1", "phase8_p2_dashboard_surface.ps1",
                      "phase8_p3_chart_surface.ps1"], others


def test_02_only_the_three_definition_only_files_are_dot_sourced() -> None:
    sourced = re.findall(r"^\.\s+\(Join-Path \$scriptDir '([^']+)'\)", _code(), re.M)
    assert sourced == [path.name for path in DOT_SOURCED], sourced


def test_03_the_ten_helpers_are_byte_identical_to_the_accepted_source() -> None:
    """COPIED, NOT PARAPHRASED, from the same file P8-1 and P8-2 copied."""
    timing = accepted._ps_code(TIMING)
    mine = _code()
    for name in COPIED_HELPERS:
        pattern = rf"^function\s+{re.escape(name)}\s*\{{.*?^\}}"
        theirs = re.search(pattern, timing, re.S | re.M)
        ours = re.search(pattern, mine, re.S | re.M)
        assert theirs and ours, name
        assert hashlib.sha256(theirs.group(0).encode()).hexdigest() == \
               hashlib.sha256(ours.group(0).encode()).hexdigest(), name


def test_04_the_reused_p82_parts_are_the_p82_ones_renamed() -> None:
    p82 = accepted._ps_code(P82)
    mine = _code()
    for name in SHARED_WITH_P82:
        original = name.replace("P83", "P82")
        theirs = re.search(rf"^function\s+{re.escape(original)}\s*\{{.*?^\}}", p82, re.S | re.M)
        ours = re.search(rf"^function\s+{re.escape(name)}\s*\{{.*?^\}}", mine, re.S | re.M)
        assert theirs and ours, name
        assert theirs.group(0).replace("P82", "P83") == ours.group(0), (
            f"{name} diverged from the accepted P8-2 {original}")


def test_05_the_runner_carries_no_powershell_6_construct() -> None:
    code = _code()
    for construct in ("??", "?.", "&&", "||", "-Parallel", "-LeafBase",
                      "-AsHashtable", "utf8NoBOM", "$IsWindows", "$PSStyle"):
        assert construct not in code, f"the runner uses {construct!r}"
    for line in code.splitlines():
        count = accepted._join_path_positional_count(line)
        assert count is None or count <= 2, f"Join-Path with >2 positionals: {line}"


def test_06_no_comma_expression_is_an_operand_of_arithmetic() -> None:
    """THE DEFECT THAT KILLED P8-1'S SECOND WINDOWS RUN."""
    import test_phase8_p1_runner_source as p81_controls
    found = p81_controls._comma_bound_arithmetic(_text())
    assert not found, f"the P8-3 runner carries comma-bound arithmetic: {found}"


# ===========================================================================
# B. EVERYTHING COMES FROM THE PROJECTION
# ===========================================================================
def test_10_no_chart_identity_anchor_type_or_range_is_spelled_in_the_runner() -> None:
    """NOT ONE ADDRESS. A hand-written range is a second declaration of
    something the manifest owns, and P7-4 is what that costs when the first one
    moves."""
    code = _code()
    for literal in re.finditer(r"['\"]\$?[A-J]\$?\d{2,4}(?::|['\"])", code):
        pytest.fail(f"the runner spells a cell address: {literal.group(0)}")
    for banned in ("Results!", "Dashboard!", "Sensitivity!$", "$B$487", "$H$279"):
        assert banned not in code, f"the runner types the reference {banned!r}"
    # THE TITLES, ANCHORS, TYPES AND RANGES ALL ARRIVE THROUGH $charts.
    # PowerShell is case-insensitive about a variable name, so the projection may
    # be reached as either the parameter $Charts or the script variable $charts.
    for projected in ("$spec.title", "$spec.anchor", "$spec.kind", "$spec.categories.range",
                      "$want.range", "$charts.bin_contract.bin_count",
                      "$charts.sensitivity_endpoint", "$charts.chart_sheet",
                      "$charts.bridge_sheet"):
        assert re.search(re.escape(projected), code, re.IGNORECASE), (
            f"the runner does not read {projected}")


def test_11_no_chart_assertion_reads_the_machine_sheets() -> None:
    """THE BRIDGE IS THE ONLY THING A CHART MAY SEE, and the only thing this
    runner compares against."""
    code = _code()
    for machine in ("_SimData", "_Calc"):
        hit = re.search(rf"(?<![A-Za-z0-9_]){re.escape(machine)}(?![A-Za-z0-9_])", code)
        # THE ONLY PERMITTED MENTION is the refusal that scans a SERIES formula
        # for it, which is the opposite of reading it.
        if hit:
            line = code[code.rfind("\n", 0, hit.start()) + 1:code.find("\n", hit.start())]
            assert "$series.Formula" in line or "foreach ($sheet in @(" in line, (
                f"the runner reads {machine}: {line.strip()!r}")
    for reader in ("Get-SimRawCell", "Get-SimBankBlock", "Get-SimSummaryValue",
                   "Get-Phase6State"):
        assert reader not in code, f"the runner reads the machine sheet via {reader}"


def test_12_the_runner_recomputes_nothing_it_checks() -> None:
    """NO SECOND HISTOGRAM, NO RECONSTRUCTED CUMULATIVE SERIES, NO RE-RANKED
    TORNADO. A runner that recomputed what it was checking would agree with
    itself about a chart that had drifted."""
    code = _code()
    for banned in ("Sort-Object", "Measure-Object -Sum", "PERCENTILE",
                   "quantile", "Group-Object", "-Descending"):
        assert banned not in code, f"the runner reimplements an analysis: {banned}"
    # A MAGNITUDE IS HOW A TORNADO IS RANKED, so absolute value is refused
    # everywhere it could stand in for an analysis. The ONE permitted use is the
    # geometric tolerance on a chart's drawn size, which is not a number the
    # model produced and which no chart series carries.
    for name, body in _functions().items():
        for hit in re.finditer(r"\[Math\]::\w+", body):
            line = body[body.rfind("\n", 0, hit.start()) + 1:body.find("\n", hit.start())]
            assert name == "Invoke-P83ChartChecks" and (
                "width_cm" in line or "height_cm" in line), (
                f"{name} reimplements an analysis: {line.strip()!r}")
    # NO BINNING: no bin width, no edge arithmetic over a min and a max. The
    # boundary matters - $widthCm is a chart's drawn size, not a bin's.
    for shape in ("$maximum - $minimum", "/ $binCount", "$width", "$lower +"):
        assert not re.search(re.escape(shape) + r"(?![A-Za-z0-9_])", code), (
            f"the runner bins for itself: {shape}")
    # THE ONE ARITHMETIC OVER PLOTTED VALUES IS COUNT CONSERVATION, and it sums
    # the frozen counts to compare against a number the SHEET published.
    total = _function("Invoke-P83HistogramChecks")
    assert "$total = $total + [double]$cell.Value" in total
    assert "run_stamp.iterations_run" in total, (
        "the expected total is not taken from the sheet")


def test_13_no_state_word_is_typed_in_the_runner() -> None:
    """THE VOCABULARY IS THE PROJECTION'S. A runner that typed CURRENT would be
    a second authority for a word the contract owns."""
    code = _code()
    for word in ("'CURRENT'", '"CURRENT"', "'STALE'", '"STALE"', "'INVALID'",
                 '"INVALID"', "'HISTORICAL'", '"HISTORICAL"', "'NOT PRODUCED'",
                 '"NOT PRODUCED"', "'OTHER Px'", '"OTHER Px"'):
        assert word not in code, f"the runner types the state word {word}"
    for projected in ("$p7.handoff.distribution_states", "$p7.handoff.profile_states",
                      "$p7.handoff.inconsistent_stamp_state",
                      "$p7.model_states.derived_status"):
        assert projected in code, f"the runner does not project {projected}"


# ===========================================================================
# C. THE OBSERVATION ORDER
# ===========================================================================
def test_20_the_orchestrator_freezes_before_it_reads_a_chart() -> None:
    """THE ORDER IS THE MEASUREMENT. Recalculate, freeze the state, freeze the
    bridge, THEN read the chart objects, then compare. A bridge cell re-read
    after the chart would be a value the chart could have influenced."""
    body = _function("Invoke-P83Observation")
    steps = ["Invoke-P83Recalculate", "Get-P83StateFrozen", "Get-P83BridgeFrozen",
             "Get-P83Charts", "Invoke-P83ChartChecks"]
    positions = [body.index(step) for step in steps]
    assert positions == sorted(positions), list(zip(steps, positions))
    for forbidden in ("$Excel.Run", "Invoke-P83Endpoint", "Set-NamedValue",
                      "Set-P83NamedText", "Set-TableCell", "Set-Phase5Fixture"):
        assert forbidden not in body, f"{forbidden} happens inside the observation"
    between = body[body.index("Get-P83StateFrozen"):body.index("Get-P83Charts")]
    assert "Calculate" not in between, "the workbook recalculates between the freezes"


def test_21_every_part_observes_through_the_orchestrator() -> None:
    for name in ("part 0", "part A1", "part A2", "part B", "part C", "part D", "part E"):
        body = _part(name)
        assert "Invoke-P83Observation" in body, f"{name} makes no observation"
        for hand_rolled in ("Get-P83Charts", "Get-P83BridgeFrozen", "Get-P83StateFrozen"):
            assert hand_rolled not in body, (
                f"{name} freezes or reads charts by hand instead of through the orchestrator")


def test_22_the_assertions_read_the_frozen_capture_not_the_sheet() -> None:
    for name in ("Invoke-P83StateChecks", "Invoke-P83QualificationChecks",
                 "Invoke-P83TornadoQualification", "Invoke-P83AnnualSeriesChecks",
                 "Invoke-P83HistogramChecks", "Compare-P83Bridge"):
        body = _function(name)
        assert "Get-P83Charts" not in body, f"{name} re-reads the chart objects"
        assert ("$Observation." in body) or ("Get-P83Frozen" in body) or (
            "$Before.Bridge" in body), name


@pytest.mark.parametrize("part", ["part B", "part D"])
def test_23_no_endpoint_runs_before_the_observation_in_the_recalc_only_parts(
        part: str) -> None:
    """THE TWO PARTS WHOSE WHOLE QUESTION IS WHAT A RECALCULATION ALONE DID."""
    body = _part(part)
    before = body[:body.index("Invoke-P83Observation")]
    for endpoint in ("Invoke-P83Endpoint", "Invoke-Phase6Simulation",
                     "Invoke-Phase5ProductionOperation", "$excel.Run("):
        assert endpoint not in before, f"{part} invokes {endpoint} before the observation"
    after = body[body.index("Invoke-P83Observation"):]
    assert "Invoke-P83Endpoint" not in after, f"{part} invokes an endpoint at all"
    if part == "part B":
        assert "$excel.Run(" not in after, "part B calls out of cell"
    else:
        assert after.count("$excel.Run(") == 1, (
            "part D takes more than the one post-observation diagnostic")
        assert after.index("$excel.Run(") > after.index("Compare-P83Bridge"), (
            "part D's diagnostic runs before its bridge comparison")


def test_24_the_transition_always_precedes_the_observation() -> None:
    transitions = {
        "part A1": "Set-Phase5Fixture",
        "part A2": "$charts.sensitivity_endpoint",
        "part B": "Set-P83NamedText",
        "part C": "Invoke-P83Endpoint",
        "part D": "Set-NamedValue -Workbook $wb -DefinedName $iterationsControl",
        "part E": "Set-TableCell",
    }
    for name, transition in transitions.items():
        body = _part(name)
        assert transition in body, f"{name} performs no transition"
        assert body.index(transition) < body.index("Invoke-P83Observation"), (
            f"{name} observes before it transitions")


# ===========================================================================
# D. THE EXCEL CHART OBJECT ITSELF
# ===========================================================================
def test_30_the_runner_reads_the_live_chart_objects_and_their_series() -> None:
    """FROM EXCEL, NOT FROM THE FILE THE BUILDER WROTE. `ChartObjects` is what
    Excel actually loaded and `Series.Formula` is the SERIES() expression it
    actually resolved - which is the only thing that can prove a range survived
    the round trip through the .xlsm."""
    body = _function("Get-P83Charts")
    for member in ("$sheet.ChartObjects()", "$object.Chart", "$chart.ChartType",
                   "$chart.SeriesCollection()", "$item.Formula", "$item.Values",
                   "$object.TopLeftCell", "$object.Width", "$object.Height",
                   "$chart.HasLegend", "$chart.ChartTitle.Text"):
        assert member in body, f"the reader never asks Excel for {member}"


def test_31_the_series_formula_is_parsed_at_the_top_level() -> None:
    """A SERIES NAME CAN CARRY A COMMA and a range cannot, so a naive split
    would tear a quoted name in half and mistake its tail for a category
    range."""
    body = _function("Get-P83SeriesParts")
    assert "$quoted" in body and "$depth" in body
    assert "-eq '\"'" in body or '"' in body


def test_32_every_projected_chart_property_is_asserted() -> None:
    body = _function("Invoke-P83ChartChecks")
    for claim in ("exactly the projected charts", "the projected Excel chart type",
                  "three-dimensional", "anchored and sized as projected",
                  "series and category range is the projected bridge range",
                  "reads the machine sheets", "equals the frozen bridge cell",
                  "draws a point where the bridge says there is none"):
        assert claim in body, f"the chart comparison does not assert {claim!r}"
    # THE THREE PERMITTED TYPES AND THE 3-D ONES ARE NAMED, not inferred.
    code = _code()
    assert "$script:P83ChartTypes = @{ 'line' = 4; 'column' = 51; 'bar' = 57 }" in code
    assert "$script:P83ThreeD" in code
    assert "xl3DColumn" in code and "xl3DBarClustered" in code


def test_33_a_no_point_is_told_apart_from_a_point_at_zero() -> None:
    """THE FABRICATED YEAR THIS WHOLE LAYER EXISTS TO PREVENT. Excel plots a
    zero-length string as ZERO; the bridge emits NA() instead, and
    `Series.Values` returns the error code for it. The runner compares the KIND
    of each point, not only its value."""
    body = _function("Test-P83NoPoint")
    assert "$script:P83ErrorCodes.ContainsKey" in body
    comparison = _function("Invoke-P83ChartChecks")
    assert "$cellBlank -ne $pointBlank" in comparison, (
        "the runner does not compare presence against presence")
    assert "fabricated" in comparison


def test_34_the_size_tolerance_is_stated_and_tight() -> None:
    """EXCEL ROUNDS A SHAPE TO WHOLE PIXELS, so the size is checked to a
    tolerance - and the tolerance is far tighter than any layout mistake."""
    body = _function("Invoke-P83ChartChecks")
    assert "-gt 0.5" in body, "the size tolerance is not 0.5 cm"
    assert "ConvertTo-P83Centimetres" in _code()


def test_35_the_freeze_pane_and_the_chart_status_placement_are_proved_live() -> None:
    body = _function("Test-P83Layout")
    assert "$window.FreezePanes" in body and "$window.SplitRow" in body
    assert "$Dashboard.freeze_panes" in body, "the expected pane is not projected"
    assert "keeps the whole Result Status block on screen" in body
    assert "above every plot" in body
    for name in ("part 0", "part E"):
        assert "Test-P83Layout" in _part(name), name


# ===========================================================================
# E. THE PARTS, AND WHAT EACH ONE IS FOR
# ===========================================================================
def test_40_the_seven_parts_are_present_and_in_order() -> None:
    code = _code()
    positions = [code.index(f"Write-P83Line '{banner}'") for banner in BANNERS]
    assert positions == sorted(positions), positions
    assert len(set(positions)) == 7


def test_41_part_a1_runs_no_sensitivity_and_proves_the_tornado_says_so() -> None:
    """A TORNADO WITH NOTHING TO PLOT MUST SAY SO. An empty chart that looks
    like a finished analysis is the failure this part exists to catch."""
    body = _part("part A1")
    assert "$charts.sensitivity_endpoint" not in body, (
        "part A1 runs sensitivity; it must not")
    assert "Invoke-P83TornadoRowChecks" in body
    assert "-Published" not in body.split("Invoke-P83TornadoRowChecks")[1].split("\n")[0], (
        "part A1 expects a published ranking")
    assert "-UnavailablePhrase $notProducedPhrase" in body


def test_42_part_a2_invokes_the_projected_sensitivity_endpoint_explicitly() -> None:
    """AN ACCEPTANCE ACTION, NOT WORKBOOK BEHAVIOUR. Nothing on the sheet starts
    an analysis to populate a chart."""
    body = _part("part A2")
    assert "Invoke-P83Endpoint -Excel $excel `\n        -Endpoint ([string]$charts.sensitivity_endpoint)" in body \
        or "-Endpoint ([string]$charts.sensitivity_endpoint)" in body
    assert "PCCM_RunSensitivity" not in _code(), (
        "the runner types the endpoint instead of projecting it")
    # A REFUSAL IS DIAGNOSED, NOT PATCHED AROUND.
    assert "does not patch around a refusal" in body
    assert "throw ('the sensitivity endpoint refused: '" in body
    assert "-Published" in body, "part A2 does not expect a published ranking"


def test_43_the_histogram_is_qualified_by_the_live_simulation_state() -> None:
    """NOT THE ANNUAL STATE, AND NOT THE PERSISTED ROW."""
    body = _function("Invoke-P83QualificationChecks")
    assert "-Key 'simulation_state'" in body
    assert "-Key 'run_stamp.simulation_status'" in body, (
        "the persisted row is not recorded beside the live one")
    assert "not by the persisted row" in body
    # THE PERSISTED ROW IS NEVER AN EXPECTATION.
    assert "Test-SimExactText -Actual $persisted" not in body
    projection = _projection()
    histogram = next(c for c in projection["charts"] if c["key"] == "histogram")
    assert histogram["state_source"] == "simulation_state"


def test_44_the_tornado_requires_both_conditions() -> None:
    body = _function("Invoke-P83TornadoQualification")
    assert "-Key 'sensitivity_availability'" in body
    assert "-Key 'simulation_state'" in body
    assert "carries the live simulation state as well" in body
    # THE SECOND CONDITION IS ITS OWN CHECK, not a clause inside the first. The
    # availability sentence compares two persisted records and is blind to a
    # model that has moved since, so a passing availability cannot carry the
    # live state with it.
    assert body.count("Add-P83Check") >= 3, (
        "the two tornado conditions are not separately recorded")
    live = body[body.index("the tornado carries the live simulation state as well"):]
    assert "Test-SimExactText -Actual $live.Value -Expected $Simulation" in live, (
        "the live condition is not decided by the live state")
    # AND WHEN THE TWO DISAGREE, the runner says so rather than accepting the
    # ranking's word for the model's state.
    assert "if ($Simulation -cne $CurrentWord) {" in body
    assert "-not (Test-SimExactText -Actual $live.Value -Expected $CurrentWord)" in body
    tornado = next(c for c in _projection()["charts"] if c["key"] == "tornado")
    assert tornado["state_source"] == "sensitivity_availability"
    assert tornado["also_qualified_by"] == "simulation_state"


@pytest.mark.parametrize("part,word", [("part D", "$calcStale"), ("part E", "$calcInvalid")])
def test_45_the_stale_and_invalid_qualifications_cannot_be_dropped(
        part: str, word: str) -> None:
    """A STALE OR INVALID CHART STAYS VISIBLE AND STAYS QUALIFIED. Preserved
    evidence is not erased; it is labelled."""
    body = _part(part)
    assert f"-Simulation {word}" in body, f"{part} does not expect {word}"
    assert "Invoke-P83QualificationChecks" in body
    assert "Invoke-P83TornadoQualification" in body
    # AND THE PRESERVED PAYLOAD IS STILL ASSERTED PRESENT.
    assert "-Published" in body, f"{part} stops expecting a published payload"
    assert "Compare-P83Bridge" in body


def test_46_the_tornado_rows_are_compared_positionally_and_signed() -> None:
    """PHASE 7 OWNS THE RANKING. Bridge row k must be Sensitivity row first+k -
    which is what "took the first N in the published order" means - and the rho
    is compared as an exact double so a negative stays negative."""
    body = _function("Invoke-P83TornadoRowChecks")
    assert "$sourceRow = $first + $index" in body
    assert "Test-SimExactDouble -Actual $bridgeRho.Value" in body
    assert "abs_rho" not in body, "the runner compares the absolute value"
    for banned in ("Sort-Object", "[Math]::Abs", "-Descending"):
        assert banned not in body, f"the runner re-ranks: {banned}"


def test_47_part_0_expects_no_point_anywhere() -> None:
    body = _part("part 0")
    assert "-ExpectedYears 0" in body
    assert "-ExpectedPx $null" in body
    assert "Invoke-P83HistogramChecks" in body
    assert "-Published" not in body, "part 0 expects a published payload"


# ===========================================================================
# F. MUTATIONS - each is a way this runner could pass while proving nothing
# ===========================================================================
def _order_ok(code: str) -> None:
    body = re.search(r"function\s+Invoke-P83Observation\s*\{(.*?)\n\}", code, re.S)
    assert body, "the orchestrator is gone"
    text = body.group(1)
    positions = [text.index(step) for step in
                 ("Invoke-P83Recalculate", "Get-P83StateFrozen", "Get-P83BridgeFrozen",
                  "Get-P83Charts", "Invoke-P83ChartChecks")]
    assert positions == sorted(positions), "the observation order is wrong"
    for forbidden in ("$Excel.Run", "Invoke-P83Endpoint"):
        assert forbidden not in text, f"{forbidden} is inside the observation"


def _oracle_ok(code: str) -> None:
    for reader in ("Get-SimRawCell", "Get-SimBankBlock", "Get-Phase6State"):
        assert reader not in code, f"the runner reads the machine sheet via {reader}"
    for banned in ("Sort-Object", "-Descending", "PERCENTILE"):
        assert banned not in code, f"the runner reimplements an analysis: {banned}"
    # A MAGNITUDE IS HOW A TORNADO IS RANKED. It is refused everywhere except the
    # one place it means a drawn size rather than a number the model produced.
    for name, body in _split_functions(code).items():
        for hit in re.finditer(r"\[Math\]::\w+", body):
            line = body[body.rfind("\n", 0, hit.start()) + 1:body.find("\n", hit.start())]
            assert name == "Invoke-P83ChartChecks" and (
                "width_cm" in line or "height_cm" in line), (
                f"{name} reimplements an analysis: {line.strip()!r}")


def _recalc_only_ok(code: str) -> None:
    for banner in ("PART B - ", "PART D - "):
        start = code.index(f"Write-P83Line '{banner}")
        body = code[start:code.index("Invoke-P83Observation", start)]
        for endpoint in ("Invoke-P83Endpoint", "Invoke-Phase6Simulation", "$excel.Run("):
            assert endpoint not in body, f"an endpoint runs before {banner}observation"


def _qualification_ok(code: str) -> None:
    live = re.search(r"function\s+Invoke-P83QualificationChecks\s*\{(.*?)\n\}", code, re.S)
    assert live, "the qualification check is gone"
    assert "-Key 'simulation_state'" in live.group(1)
    assert "Test-SimExactText -Actual $persisted" not in live.group(1), (
        "the persisted row became an expectation")
    tornado = re.search(r"function\s+Invoke-P83TornadoQualification\s*\{(.*?)\n\}", code, re.S)
    assert tornado, "the tornado qualification is gone"
    assert "-Key 'simulation_state'" in tornado.group(1), (
        "the tornado stopped asking whether the run matches the model")


def _fabrication_ok(code: str) -> None:
    body = re.search(r"function\s+Invoke-P83ChartChecks\s*\{(.*?)\n\}", code, re.S)
    assert body, "the chart comparison is gone"
    assert "$cellBlank -ne $pointBlank" in body.group(1), (
        "a drawn point is no longer compared against an absent cell")


RULES = (_order_ok, _oracle_ok, _recalc_only_ok, _qualification_ok, _fabrication_ok)


@pytest.mark.parametrize("name,mutate", [
    # THE HISTOGRAM QUALIFIED FROM THE PERSISTED ROW - the defect the pre-Windows
    # correction existed to remove, reintroduced in the harness.
    ("the histogram is qualified from the persisted status row",
     lambda code: code.replace(
         "    $live = Get-P83Frozen -Observation $Observation -Key 'simulation_state'\n"
         "    $persisted = Get-P83Frozen -Observation $Observation -Key 'run_stamp.simulation_status'",
         "    $live = Get-P83Frozen -Observation $Observation -Key 'run_stamp.simulation_status'\n"
         "    $persisted = Get-P83Frozen -Observation $Observation -Key 'run_stamp.simulation_status'", 1)),
    # THE TORNADO IGNORING THE LIVE STATE - one condition where there are two.
    ("the tornado stops asking whether the run matches the model",
     lambda code: code.replace(
         "    $live = Get-P83Frozen -Observation $Observation -Key 'simulation_state'\n"
         "    $text = [string]$availability.Value",
         "    $text = [string]$availability.Value", 1)),
    # THE CHARTS READ BEFORE THE CELLS THEY ARE COMPARED AGAINST.
    ("the chart objects are read before the bridge is frozen",
     lambda code: code.replace(
         "    $bridge = Get-P83BridgeFrozen -Workbook $Workbook -Charts $Charts\n"
         "    $objects = Get-P83Charts -Workbook $Workbook -SheetName ([string]$Charts.chart_sheet)",
         "    $objects = Get-P83Charts -Workbook $Workbook -SheetName ([string]$Charts.chart_sheet)\n"
         "    $bridge = Get-P83BridgeFrozen -Workbook $Workbook -Charts $Charts", 1)),
    # AN ENDPOINT BEFORE PART D's OBSERVATION.
    ("part D runs an endpoint before observing",
     lambda code: code.replace(
         "    $observationD = Invoke-P83Observation",
         "    $null = Invoke-P83Endpoint -Excel $excel -Endpoint 'PCCM_Calculate'\n"
         "    $observationD = Invoke-P83Observation", 1)),
    # A DRAWN ZERO ACCEPTED WHERE THE BRIDGE SAYS THERE IS NO POINT.
    ("a fabricated point stops being refused",
     lambda code: code.replace("                if ($cellBlank -ne $pointBlank) {",
                               "                if ($false) {", 1)),
    # THE RUNNER RANKING BY MAGNITUDE FOR ITSELF. A tornado IS a magnitude
    # ranking, so an absolute value here would be a second analytical authority -
    # which is exactly what the geometric tolerance in Invoke-P83ChartChecks is
    # not, and why that rule is scoped to a function rather than dropped.
    ("the runner ranks the drivers by magnitude",
     lambda code: code.replace(
         "    $mismatched = New-Object System.Collections.ArrayList\n"
         "    $plotted = 0",
         "    $mismatched = New-Object System.Collections.ArrayList\n"
         "    $plotted = 0\n"
         "    $largest = [Math]::Abs([double]$rhos[0].Value)", 1)),
    # THE RUNNER SORTING THE TORNADO FOR ITSELF.
    ("the runner re-ranks the drivers",
     lambda code: code.replace(
         "    $mismatched = New-Object System.Collections.ArrayList",
         "    $mismatched = New-Object System.Collections.ArrayList\n"
         "    $names = @($names | Sort-Object)", 1)),
    # THE MACHINE SHEET READ DIRECTLY.
    ("the runner reads the machine sheet",
     lambda code: code.replace(
         "function Get-P83BridgeFrozen {",
         "function Get-P83BridgeFrozen {\n    $null = Get-SimRawCell", 1)),
    # AN OUT-OF-CELL CALL MOVED INSIDE THE OBSERVATION.
    ("an out-of-cell call moves inside the observation",
     lambda code: code.replace(
         "    $state = Get-P83StateFrozen -Workbook $Workbook -P8 $P8 -Charts $Charts",
         "    $null = $Excel.Run('PCCM_CalculationStatus')\n"
         "    $state = Get-P83StateFrozen -Workbook $Workbook -P8 $P8 -Charts $Charts", 1)),
])
def test_50_each_way_of_passing_while_proving_nothing_is_refused(
        name: str, mutate) -> None:
    """EACH MUTATION APPLIED TO A COPY OF THE SOURCE IN MEMORY, with the rules
    above re-run over it. Nothing on disk changes."""
    mutated = mutate(_code())
    assert mutated != _code(), f"the mutation '{name}' changed nothing"
    refused = []
    for rule in RULES:
        try:
            rule(mutated)
        except (AssertionError, ValueError) as failure:
            refused.append(f"{rule.__name__}: {failure}")
    assert refused, f"'{name}' survived every rule"


def test_51_the_rules_pass_on_the_unmutated_runner() -> None:
    """SO THE NINE REFUSALS ABOVE ARE REFUSALS OF THE MUTATION, not of the
    fixture."""
    code = _code()
    for rule in RULES:
        rule(code)
    projection = _projection()
    assert len(projection["charts"]) == 4
    assert projection["sensitivity_endpoint"].startswith("PCCM_Run")
