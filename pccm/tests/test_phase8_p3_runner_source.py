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
import yaml  # noqa: E402

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
SPEC = PCCM_ROOT / "spec"

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


def _part_of(code: str, name: str) -> str:
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


def _part(name: str) -> str:
    return _part_of(_code(), name)


def _gate_b_cases() -> dict:
    if "gateb" not in _CACHE:
        _CACHE["gateb"] = json.loads(
            (BUILD / "phase6_gate_b_cases.json").read_text(encoding="utf-8"))
    return _CACHE["gateb"]


def _dashboard() -> dict:
    if "dash" not in _CACHE:
        _CACHE["dash"] = json.loads(
            (BUILD / "phase8_dashboard_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["dash"]


def _sim_state_word(variable: str) -> str:
    """The word a `$sim*` runner variable actually resolves to.

    The runner reads the SIMULATION axis out of the gate-B projection, which
    reads it from `sim_contract.yaml: label_sets.sim_state`. Resolving the index
    here means a control compares against the CONTRACT, not against a name a
    runner chose - the distinction that matters, because the calculation axis
    spells two of its four states identically.
    """
    match = re.search(
        re.escape(variable) + r"\s*=\s*\[string\]\$gateBCases\.vocabulary\.sim_states\[(\d+)\]",
        _code())
    assert match, f"{variable} is not read from the projected simulation axis"
    states = _gate_b_cases()["vocabulary"]["sim_states"]
    return str(states[int(match.group(1))])


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


@pytest.mark.parametrize("part,word,expected", [
    ("part D", "$simStale", "STALE"), ("part E", "$simInvalid", "INVALID")])
def test_45_the_stale_and_invalid_qualifications_cannot_be_dropped(
        part: str, word: str, expected: str) -> None:
    """A STALE OR INVALID CHART STAYS VISIBLE AND STAYS QUALIFIED. Preserved
    evidence is not erased; it is labelled.

    AND THE WORD IS THE SIMULATION AXIS'S. Naming a variable is not enough - the
    control resolves it through the projection to the contract's own label, so a
    variable that pointed at the calculation axis would be caught even where the
    two axes happen to spell a state the same way."""
    body = _part(part)
    assert f"-Simulation {word}" in body, f"{part} does not expect {word}"
    assert _sim_state_word(word) == expected, (
        f"{part}'s {word} does not resolve to {expected} on the simulation axis")
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
# E2. THE TWO AXES, AND THE ONE PROPERTY THE FIRST WINDOWS RUN COULD NOT FIND
# ===========================================================================
# WHAT THE FIRST LIVE RUN ESTABLISHED, and it is worth stating because both
# defects below were the runner's and neither was production's.
#
# The calculation axis is NOT CALCULATED / CURRENT / STALE / INVALID. The
# simulation axis is CURRENT / STALE / INVALID and nothing else - sim_contract
# says so in as many words, and a workbook with no publication holds a BLANK
# status rather than a fourth label. The runner read all four calculation words
# and asserted one of them, NOT CALCULATED, of the LIVE SIMULATION state cell.
# The accepted owner returned INVALID, which is what its first ordered rule
# says an untouched workbook gets, and the runner called that a failure.
def test_60_the_simulation_axis_has_no_fourth_state() -> None:
    """SO `NOT CALCULATED` CANNOT BE A SIMULATION EXPECTATION - it is not a word
    on that axis at all."""
    states = _gate_b_cases()["vocabulary"]["sim_states"]
    assert states == ["CURRENT", "STALE", "INVALID"], states
    assert "NOT CALCULATED" not in states
    # AND IT IS THE CALCULATION AXIS THAT OWNS THAT WORD.
    calc = json.loads((BUILD / "phase7_acceptance_inspection.json").read_text(
        encoding="utf-8"))["model_states"]["derived_status"]
    assert calc[0] == "NOT CALCULATED"
    assert set(states) < set(calc), (
        "the two axes no longer overlap the way these controls assume")


def test_61_part_0_expects_the_state_the_accepted_owner_derives() -> None:
    """INVALID, AND FROM THE CONTRACT RATHER THAN FROM A GUESS.

    sim_contract.yaml's derivation is ORDERED and the first matching rule wins.
    Rule 1 is `current_prerequisites_do_not_resolve -> INVALID`; rule 2 is
    `no_successful_snapshot_exists -> null`. On an untouched workbook the
    calculation is NOT CALCULATED, so the simulation's prerequisites do not
    resolve and rule 1 fires before rule 2 is ever reached."""
    contract = yaml.safe_load((SPEC / "sim_contract.yaml").read_text(encoding="utf-8"))
    rules = contract["sim_state"]["derivation"]["rules"]
    assert contract["sim_state"]["derivation"]["ordered"] is True
    first = min(rules, key=lambda rule: rule["order"])
    assert first["condition"] == "current_prerequisites_do_not_resolve"
    assert first["status"] == "INVALID"
    assert contract["sim_state"]["definitions"]["INVALID"] == (
        "current simulation prerequisites do not resolve")
    # AND THE RUNNER EXPECTS EXACTLY THAT.
    body = _part("part 0")
    assert "-Simulation $simInvalid" in body
    assert _sim_state_word("$simInvalid") == "INVALID"


def test_62_no_calculation_word_is_asserted_of_a_simulation_cell() -> None:
    """THE AXES STAY APART. Every `-Simulation` and `-CurrentWord` argument comes
    off the simulation axis; the calculation axis survives only where the cell
    being compared is PCCM_CalculationStatus."""
    code = _code()
    for match in re.finditer(r"-(?:Simulation|CurrentWord)\s+(\$\w+)", code):
        assert match.group(1).startswith("$sim"), (
            f"a simulation expectation reads {match.group(1)}")
    for calc in re.finditer(r"\$calc(?:Current|Stale|Invalid|NotCalculated)\b", code):
        line = code[code.rfind("\n", 0, calc.start()) + 1:code.find("\n", calc.start())]
        assert ("model_states.derived_status" in line or
                "PCCM_CalculationStatus" in line or
                "$calcStatus" in line), (
            f"a calculation word is used away from the calculation axis: {line.strip()!r}")
    assert "$calcNotCalculated" not in code, (
        "the runner still carries the calculation axis's fourth word")


def test_63_a_qualification_check_asserts_the_state_it_names() -> None:
    """THE VACUOUS CHECK, REFUSED. It printed `qualified NOT CALCULATED` beside a
    cell reading INVALID and passed, because it asserted only `not CURRENT` -
    true of every word that is not CURRENT, including one from the wrong axis."""
    checked = 0
    for name in ("Invoke-P83StateChecks", "Invoke-P83QualificationChecks",
                 "Invoke-P83TornadoQualification"):
        body = _function(name)
        calls = [match.start() for match in re.finditer(r"Add-P83Check\b", body)]
        for index, start in enumerate(calls):
            end = calls[index + 1] if index + 1 < len(calls) else len(body)
            call = body[start:end]
            title = call[:call.find("\n")]
            # WHICH STATE VARIABLE THIS CHECK'S NAME PROMISES. Only the title
            # line is read: a state word interpolated into the sentence a reader
            # sees is the claim the condition then has to make good.
            promised = [word for word in re.findall(r"\$\w+", title)
                        if word in ("$Simulation", "$Distribution", "$Profile")]
            if not promised:
                continue
            checked += 1
            # AND THE CONDITION - this call's own text, never the next one's -
            # must compare the observed cell against that very variable.
            condition = call[len(title):]
            for word in promised:
                assert f"-Expected {word}" in condition, (
                    f"{name} names {word} and does not assert it: "
                    f"{title.strip()[:90]!r}")
    assert checked >= 4, (
        f"only {checked} state-naming checks were examined; the sweep has stopped "
        "finding them")


def test_64_the_publication_question_is_asked_without_the_state_word() -> None:
    """AN INVALID MODEL IS NOT A PUBLISHED RUN, AND NOT AN UNPUBLISHED ONE
    EITHER. The contract keeps the two apart - the no-snapshot rule returns a
    BLANK, not a word - so the runner must ask publication directly."""
    body = _function("Invoke-P83PublicationExistence")
    assert "run_stamp.run_id" in body
    assert "Test-P83Blank -Cell $run" in body
    assert "distribution.count" in body
    for leak in ("simulation_state", "$Simulation", "Test-SimExactText"):
        assert leak not in body, (
            f"the publication question consults the state word via {leak}")
    part0 = _part("part 0")
    assert "Invoke-P83PublicationExistence -Observation $observation0" in part0
    # AND THE SENSITIVITY SENTENCE SAYS THE SAME THING SEPARATELY.
    assert "-UnavailablePhrase 'No simulation has been published'" in part0


# ---------------------------------------------------------------------------
# THE FREEZE THAT WAS NOT THERE
# ---------------------------------------------------------------------------
# The runner asked the DASHBOARD PROJECTION for `freeze_panes`. That projection
# is built from `phase6_shell.dashboard`, and the declaration lives on the SHEET
# - `sheets[Dashboard].freeze_panes`, a worksheet-layout property every sheet
# has. The projection never carried it, StrictMode 2.0 refuses a property that
# is not there, and the session aborted. The fix carries the field; it does not
# type the cell.
def test_65_the_freeze_is_declared_once_and_projected_from_there() -> None:
    manifest = yaml.safe_load((SPEC / "workbook.yaml").read_text(encoding="utf-8"))
    dashboard = manifest["phase6_shell"]["dashboard"]
    assert "freeze_panes" not in dashboard, (
        "the shell block now declares a freeze too; there would be two authorities")
    sheet = next(s for s in manifest["sheets"] if s["name"] == dashboard["sheet"])
    declared = sheet["freeze_panes"]
    assert declared, "the Dashboard sheet declares no freeze"
    assert _dashboard()["freeze_panes"] == declared, (
        "the projection does not carry the sheet's declaration")


def test_66_the_projected_freeze_exists_before_windows() -> None:
    """THE POINT OF THIS ONE. A missing property is not a failed assertion on
    Windows - it is an aborted session, forty minutes in, with the rest of the
    run unobserved."""
    projection = _dashboard()
    assert "freeze_panes" in projection, (
        "the runner would abort on a property that is not there")
    assert re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", str(projection["freeze_panes"]))
    # AND THE FREEZE IS ABOVE THE CHARTS, or it qualifies nothing.
    row = int(re.sub(r"[A-Z]", "", str(projection["freeze_panes"])))
    assert row <= int(projection["chart_region"]["first_row"])


def test_67_the_runner_reads_the_freeze_from_the_projection_only() -> None:
    code = _code()
    assert "$Dashboard.freeze_panes" in code, (
        "the runner does not read the projected declaration")
    for literal in ("'A17'", '"A17"'):
        assert literal not in code, f"the runner types the freeze cell {literal}"
    # NO OTHER OBJECT MAY ANSWER THIS QUESTION.
    for wrong in ("$Charts.freeze_panes", "$charts.freeze_panes", "$P8.freeze_panes",
                  "$p8.freeze_panes", "$manifest.freeze_panes"):
        assert wrong not in code, f"the freeze is read off {wrong}"


def test_68_the_runner_converts_the_cell_to_a_split() -> None:
    """EXCEL REPORTS A SPLIT, NOT A CELL. Everything above and left of the
    declared cell is frozen, so both coordinates convert - and the conversion is
    general, because the runner does not get to know which cell it is handed."""
    body = _function("Test-P83Layout")
    assert "$window.SplitRow" in body and "$window.SplitColumn" in body
    assert "$window.FreezePanes" in body
    assert "ConvertTo-P83ColumnNumber -Letters $wantLetters" in body, (
        "the column half of the declared cell is not converted")
    assert "($rows -eq ($wantRow - 1))" in body
    assert "($columns -eq ($wantColumn - 1))" in body, (
        "the column split is compared against a constant rather than the "
        "declaration")


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
    # AND THE DIVERGENCE CHECK ASSERTS THE STATE IT NAMES. Asserting only
    # `not CURRENT` passes for every other word, including one off the wrong
    # axis, which is how it printed NOT CALCULATED beside a cell reading INVALID.
    divergence = tornado.group(1)[tornado.group(1).find(
        "even though its ranking still names the published run"):]
    assert "Test-SimExactText -Actual $live.Value -Expected $Simulation" in divergence, (
        "the divergence check names a state it does not assert")


def _fabrication_ok(code: str) -> None:
    body = re.search(r"function\s+Invoke-P83ChartChecks\s*\{(.*?)\n\}", code, re.S)
    assert body, "the chart comparison is gone"
    assert "$cellBlank -ne $pointBlank" in body.group(1), (
        "a drawn point is no longer compared against an absent cell")


def _axis_ok(code: str) -> None:
    """EVERY SIMULATION EXPECTATION RESOLVES, THROUGH THE PROJECTION, TO THE WORD
    THE CONTRACT REQUIRES FOR THAT PART. Names prove nothing: the calculation
    axis spells two of its four states exactly as the simulation axis does, so
    the rule follows the index into the contract's own label set."""
    states = _gate_b_cases()["vocabulary"]["sim_states"]
    resolved: dict[str, str] = {}
    for match in re.finditer(
            r"(\$\w+)\s*=\s*\[string\]\$gateBCases\.vocabulary\.sim_states\[(\d+)\]",
            code):
        resolved[match.group(1)] = str(states[int(match.group(2))])
    for part, expected in (("part 0", "INVALID"), ("part D", "STALE"),
                           ("part E", "INVALID")):
        body = _part_of(code, part)
        found = re.search(r"-Simulation\s+(\$\w+)", body)
        assert found, f"{part} makes no simulation-state expectation"
        word = resolved.get(found.group(1))
        assert word is not None, (
            f"{part} expects {found.group(1)}, which is not read off the "
            "projected simulation axis")
        assert word == expected, f"{part} expects {word}, not {expected}"
    assert "$calcNotCalculated" not in code, (
        "the calculation axis's fourth word is back in the runner")


def _freeze_ok(code: str) -> None:
    """THE FREEZE COMES FROM THE PROJECTED SHEET DECLARATION AND IS CONVERTED."""
    assert "$Dashboard.freeze_panes" in code, "the projected declaration is not read"
    for literal in ("'A17'", '"A17"'):
        assert literal not in code, f"the runner types the freeze cell {literal}"
    for wrong in ("$Charts.freeze_panes", "$charts.freeze_panes", "$P8.freeze_panes",
                  "$p8.freeze_panes", "$manifest.freeze_panes"):
        assert wrong not in code, f"the freeze is read off {wrong}"
    body = re.search(r"function\s+Test-P83Layout\s*\{(.*?)\n\}", code, re.S)
    assert body, "the layout check is gone"
    assert "ConvertTo-P83ColumnNumber -Letters $wantLetters" in body.group(1)
    assert "($columns -eq ($wantColumn - 1))" in body.group(1), (
        "the column split is not compared against the declaration")


RULES = (_order_ok, _oracle_ok, _recalc_only_ok, _qualification_ok, _fabrication_ok,
         _axis_ok, _freeze_ok)


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
    # ---- THE SIX FROM THE FIRST WINDOWS RUN ----
    # PART 0 EXPECTING THE WRONG STATE. The runner did exactly this: it asserted
    # the calculation axis's NOT CALCULATED of the live SIMULATION state cell.
    ("part 0 expects the wrong simulation state",
     lambda code: code.replace(
         "        -Simulation $simInvalid -Distribution $notProduced",
         "        -Simulation $simCurrent -Distribution $notProduced", 1)),
    # THE SAME DEFECT IN ITS ORIGINAL FORM: a word from the other axis.
    ("part 0 expects a calculation word of a simulation cell",
     lambda code: code.replace(
         "$simCurrent = [string]$gateBCases.vocabulary.sim_states[0]",
         "$calcNotCalculated = [string]$p7.model_states.derived_status[0]\n"
         "$simCurrent = [string]$gateBCases.vocabulary.sim_states[0]", 1).replace(
         "        -Simulation $simInvalid -Distribution $notProduced",
         "        -Simulation $calcNotCalculated -Distribution $notProduced", 1)),
    # A QUALIFICATION CHECK THAT STOPS COMPARING THE STATE - the vacuous check,
    # restored exactly as it passed on Windows while printing the wrong word.
    ("a qualification check stops comparing the state it names",
     lambda code: code.replace(
         "            ((Test-SimExactText -Actual $live.Value -Expected $Simulation) -and `\n"
         "             (-not (Test-SimExactText -Actual $live.Value -Expected $CurrentWord))) `",
         "            (-not (Test-SimExactText -Actual $live.Value -Expected $CurrentWord)) `", 1)),
    # THE FREEZE CELL TYPED INTO THE RUNNER.
    ("the freeze cell is hard-coded",
     lambda code: code.replace(
         "        $wantCell = [string]$Dashboard.freeze_panes",
         "        $wantCell = 'A17'", 1)),
    # THE FREEZE READ OFF THE WRONG PROJECTION OBJECT - which is the shape of the
    # original defect, and StrictMode turns it into an aborted session.
    ("the freeze is read off the charts projection",
     lambda code: code.replace(
         "        $wantCell = [string]$Dashboard.freeze_panes",
         "        $wantCell = [string]$Charts.freeze_panes", 1)),
    # PART D'S STALE QUALIFICATION QUIETLY BECOMING CURRENT.
    ("part D stops expecting STALE",
     lambda code: code.replace(
         "        -Simulation $simStale -Distribution $historical",
         "        -Simulation $simCurrent -Distribution $historical", 1)),
    # AND PART E'S INVALID.
    ("part E stops expecting INVALID",
     lambda code: code.replace(
         "        -Simulation $simInvalid -Distribution $historical",
         "        -Simulation $simCurrent -Distribution $historical", 1)),
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
    """SO THE SIXTEEN REFUSALS ABOVE ARE REFUSALS OF THE MUTATION, not of the
    fixture."""
    code = _code()
    for rule in RULES:
        rule(code)
    projection = _projection()
    assert len(projection["charts"]) == 4
    assert projection["sensitivity_endpoint"].startswith("PCCM_Run")
