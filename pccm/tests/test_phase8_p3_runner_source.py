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
    "Format-P83Cell", "Test-P83Blank", "Invoke-P83Recalculate",
)

# THE ONE PART THAT IS DELIBERATELY NOT P8-2's ANY MORE, and why.
#
# P8-2 mirrored scalars: a Dashboard cell against the Results cell behind it,
# neither of which is ever an error in a healthy workbook, so answering
# "not equal" whenever either side was an error cost that runner nothing.
#
# P8-3 COMPARES CHART BRIDGES, and those carry NA() by design after the
# published window - a bar that does not exist must not be drawn as a zero. The
# inherited rule made every preservation check over a partly-empty block report
# `<Int32 -2146826246> -> <Int32 -2146826246>`: the same value, called movement,
# six times in Run 3. The divergence is declared here rather than hidden, and
# the control below requires it to be a real one.
DIVERGED_FROM_P82 = {
    "Test-P83SameCellValue": (
        "Excel error values need semantic equality: the same error code is the "
        "same absence, a different code is a broken reference, and neither is "
        "blank or zero."),
}

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
    """REUSING A LIFECYCLE THAT HAS RUN ON WINDOWS MEANS REUSING IT - and where
    P8-3 needs different behaviour, saying so rather than editing quietly."""
    p82 = accepted._ps_code(P82)
    mine = _code()

    def bodies(name: str) -> tuple[str, str]:
        original = name.replace("P83", "P82")
        theirs = re.search(rf"^function\s+{re.escape(original)}\s*\{{.*?^\}}", p82, re.S | re.M)
        ours = re.search(rf"^function\s+{re.escape(name)}\s*\{{.*?^\}}", mine, re.S | re.M)
        assert theirs and ours, name
        return theirs.group(0).replace("P82", "P83"), ours.group(0)

    for name in SHARED_WITH_P82:
        theirs, ours = bodies(name)
        assert theirs == ours, (
            f"{name} diverged from the accepted P8-2 original without being "
            "declared in DIVERGED_FROM_P82")
    # A DECLARED DIVERGENCE MUST BE A REAL ONE. Declaring a part changed and
    # leaving it identical would retire a control by paperwork.
    for name, reason in DIVERGED_FROM_P82.items():
        theirs, ours = bodies(name)
        assert theirs != ours, f"{name} is declared changed and is not"
        assert len(reason) > 40, f"{name} is declared changed with no reason"
    # AND NOTHING IS BOTH.
    assert not (set(SHARED_WITH_P82) & set(DIVERGED_FROM_P82))


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
    is compared through the one semantic owner, which applies no tolerance, so a
    negative stays negative and a changed magnitude of any size is a change."""
    body = _function("Invoke-P83TornadoRowChecks")
    assert "$sourceRow = $first + $index" in body
    assert "Test-P83SameCellValue -Left $bridgeRho -Right $sourceRho" in body
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
# E3. EVERY PROJECTION PROPERTY THE RUNNER TOUCHES
# ===========================================================================
# TWO WINDOWS RUNS, TWO MISSING PROPERTIES. `freeze_panes` was read off the
# dashboard projection, which does not own it; `columns` and `first_row` were
# read off the Phase-6 gate-B inspection, which describes the machine sheet and
# the publication banks and has no Sensitivity presentation layout at all.
#
# NEITHER WAS A FAILED ASSERTION. StrictMode 2.0 makes an absent property a
# terminated session, so each cost a full Windows run and left everything after
# the crash unobserved. One missing-property crash per run is not a rate anyone
# should accept, and the fix is not to be more careful - it is to resolve every
# dereference here, against the artefacts the runner will actually be handed.

# WHICH LOADED VARIABLE IS WHICH ARTEFACT, and the parameter names the same
# objects arrive under inside the functions.
PROJECTION_ROOTS = {
    "manifest": "stage_b_manifest.json",
    "inspection": "phase5_gate_b_inspection.json",
    "simInspection": "phase6_gate_b_inspection.json",
    "gateBCases": "phase6_gate_b_cases.json",
    "p7": "phase7_acceptance_inspection.json",
    "cases": "phase7_acceptance_cases.json",
    "p8": "phase8_results_inspection.json",
    "dash": "phase8_dashboard_inspection.json",
    "charts": "phase8_charts_inspection.json",
    "P8": "phase8_results_inspection.json",
    "Dashboard": "phase8_dashboard_inspection.json",
    "Charts": "phase8_charts_inspection.json",
    "P7": "phase7_acceptance_inspection.json",
}

# POWERSHELL'S OWN MEMBERS, not keys of the JSON. A trailing .Count on a
# projected list is the language's, and reflection through .PSObject is how the
# runner enumerates a block whose field names the manifest owns.
PS_MEMBERS = ("Count", "Keys", "Values", "Length", "ToCharArray", "PSObject",
              "Properties", "Name")


def _resolve_projection(obj, parts: list[str]) -> str | None:
    """Walk a dotted path. A list resolves through its first element, because a
    projected list is homogeneous by construction and the runner indexes it."""
    for part in parts:
        while isinstance(obj, list):
            if not obj:
                return f"empty list before .{part}"
            obj = obj[0]
        if not isinstance(obj, dict):
            return f"not an object at .{part}"
        if part not in obj:
            return f"no key {part!r} (has: {', '.join(sorted(obj))})"
        obj = obj[part]
    return None


def test_70_every_projection_property_the_runner_reads_exists() -> None:
    """THE CONTROL THAT REPLACES A WINDOWS RUN. Nothing here needs Excel: the
    generated artefacts are on disk and every path the runner walks can be
    walked now."""
    data = {name: json.loads((BUILD / filename).read_text(encoding="utf-8"))
            for name, filename in PROJECTION_ROOTS.items()}
    code = _code()
    pattern = r"\$(" + "|".join(PROJECTION_ROOTS) + r")((?:\.[A-Za-z_][A-Za-z0-9_]*)+)"
    unresolved: list[str] = []
    checked = 0
    for match in re.finditer(pattern, code):
        name, chain = match.group(1), match.group(2)
        parts = chain.strip(".").split(".")
        while parts and parts[-1] in PS_MEMBERS:
            parts = parts[:-1]
        if not parts:
            continue
        checked += 1
        why = _resolve_projection(data[name], parts)
        if why:
            unresolved.append(f"${name}{chain} -> {why}")
    assert not unresolved, (
        "the runner would abort on a property that is not there:\n  " +
        "\n  ".join(sorted(set(unresolved))))
    assert checked >= 50, (
        f"only {checked} dereferences were resolved; the sweep has stopped "
        "finding them and would pass a runner that read nothing")


def test_71_the_property_audit_catches_a_property_that_is_not_there() -> None:
    """SO test_70 IS NOT VACUOUS. Both real failures are replayed against it."""
    data = {name: json.loads((BUILD / filename).read_text(encoding="utf-8"))
            for name, filename in PROJECTION_ROOTS.items()}
    # RUN 1: the freeze read off an object that does not own it.
    assert _resolve_projection(data["charts"], ["freeze_panes"]) is not None
    # RUN 2: the Sensitivity layout read off the Phase-6 machine inspection.
    assert _resolve_projection(data["simInspection"], ["columns"]) is not None
    assert _resolve_projection(data["simInspection"], ["first_row"]) is not None
    # AND THE CORRECT OWNERS RESOLVE.
    assert _resolve_projection(data["dash"], ["freeze_panes"]) is None
    assert _resolve_projection(
        data["charts"], ["sensitivity_source", "columns", "column"]) is None
    assert _resolve_projection(
        data["charts"], ["sensitivity_source", "first_row"]) is None


# ---------------------------------------------------------------------------
# THE TORNADO'S TWO SIDES
# ---------------------------------------------------------------------------
def test_72_the_tornado_source_is_owned_by_the_sensitivity_block() -> None:
    """PHASE-6 SHELL DECLARES THE SENSITIVITY SHEET'S LAYOUT; the chart
    projection carries only the two fields the tornado mirrors, looked up BY KEY
    so a neighbouring column cannot be picked up by position - `abs_rho` sits
    directly beside `rho`."""
    manifest = yaml.safe_load((SPEC / "workbook.yaml").read_text(encoding="utf-8"))
    declared = {str(column["key"]): str(column["column"])
                for column in manifest["phase6_shell"]["sensitivity"]["columns"]}
    source = _projection()["sensitivity_source"]
    assert source["first_row"] == manifest["phase6_shell"]["sensitivity"]["first_row"]
    for column in source["columns"]:
        assert declared[column["key"]] == column["column"], (
            f"{column['key']} is projected at {column['column']}, declared at "
            f"{declared[column['key']]}")
    # THE SIGNED COLUMN, NOT THE MAGNITUDE.
    rho = next(c for c in source["columns"] if c["key"] == "rho")
    assert rho["column"] == declared["rho"] != declared["abs_rho"], (
        "the tornado mirrors the absolute magnitude; the sign is the point")
    # AND EXACTLY WHAT THE BRIDGE PLOTS, in the same order.
    plotted = [c["key"] for c in _projection()["bridge"]["drivers"]["columns"]]
    assert [c["key"] for c in source["columns"]] == plotted


def test_73_the_runner_reads_the_sensitivity_layout_from_that_owner_only() -> None:
    code = _code()
    assert "-Sensitivity $charts.sensitivity_source" in code, (
        "the tornado does not read the projected sensitivity source")
    # NOT FROM AN OBJECT THAT MERELY HAS A `columns` PROPERTY.
    for wrong in ("$simInspection.columns", "$simInspection.first_row",
                  "-Sensitivity $simInspection", "-Sensitivity $p8",
                  "-Sensitivity $dash", "-Sensitivity $manifest"):
        assert wrong not in code, f"the tornado source is read off {wrong}"
    # AND NO SENSITIVITY ADDRESS IS TYPED. The sheet's data starts at D13/E13;
    # neither the columns nor the row may appear as a literal.
    body = _function("Invoke-P83TornadoRowChecks")
    for typed in ("'D'", '"D"', "'E'", '"E"', "'D13'", "'E13'", "13"):
        assert typed not in body, f"the tornado types the sensitivity address {typed}"
    assert "$columns['driver_name']" in body and "$columns['rho']" in body
    assert "$columns['abs_rho']" not in body, (
        "the tornado reads the absolute magnitude")


def test_74_the_tornado_preserves_the_published_order_and_the_sign() -> None:
    """PHASE 7 RANKED THEM. This runner checks that bridge row k is Sensitivity
    row first+k and nothing else - it does not re-derive the order, and it does
    not lose a negative driver to a magnitude comparison."""
    body = _function("Invoke-P83TornadoRowChecks")
    assert "$sourceRow = $first + $index" in body, (
        "the source row is not the positional one")
    assert "for ($index = 0; $index -lt $names.Count; $index++)" in body, (
        "the plotted rows are not walked in published order")
    assert "Test-P83SameCellValue -Left $bridgeRho -Right $sourceRho" in body, (
        "the signed rho is not compared through the one semantic owner")
    for banned in ("Sort-Object", "-Descending", "[Math]::Abs", "$names.Count - 1",
                   "$index--", "[array]::Reverse"):
        assert banned not in body, f"the tornado reorders for itself: {banned}"
    # N IS THE PROJECTION'S, and the plotted count may not exceed it.
    assert "$plotted -le $names.Count" in body
    assert _projection()["bridge"]["drivers"]["row_count"] == 10
    # AND NOTHING IS FABRICATED WHEN FEWER THAN N ARE PUBLISHED.
    assert "if ($sourcePresent -ne $bridgePresent) {" in body
    assert "if (-not $sourcePresent) {" in body


def test_75_a_missing_source_column_stops_the_runner_rather_than_misreading() -> None:
    """AN ABSENT KEY WOULD BUILD THE ADDRESS `13` AND READ SOME OTHER CELL. The
    runner refuses instead, which is the difference between a wrong answer and
    no answer."""
    body = _function("Invoke-P83TornadoRowChecks")
    assert "foreach ($required in @('driver_name', 'rho')) {" in body
    assert "if (-not $columns.ContainsKey($required)) {" in body
    assert "throw (" in body


# ===========================================================================
# E4. THE COMPARATOR, EXECUTED
# ===========================================================================
# WHAT RUN 3 EXPOSED. 187 checks, 7 failed, and six of the seven reported pairs
# like `<Int32 -2146826246> -> <Int32 -2146826246>` - the SAME value, declared
# different. -2146826246 is Excel's #N/A through COM, which is exactly what the
# chart bridge emits after the published window so a bar that does not exist is
# not drawn as a zero. The old equality answered `not equal` the moment either
# side was an error, before ever reading what the error was.
#
# THE SEVENTH had the same shape one type along: the tornado's name comparison
# went through a text helper that returns false when the actual value is not a
# String, so a numeric label was refused before any comparison happened.
#
# THESE CONTROLS RUN THE COMPARATOR. It is pure logic with no COM in it, so its
# behaviour can be established here rather than asserted about its source - and
# a truth table is the only honest way to show that an equality distinguishes
# what it must and nothing more.
PWSH = "/opt/pwsh/pwsh"

TRUTH_TABLE = [
    # (label, left, right, expected)
    ("same #N/A",              "(C $na '#N/A')",   "(C $na '#N/A')",   True),
    ("#N/A vs #REF!",          "(C $na '#N/A')",   "(C $ref '#REF!')", False),
    ("#N/A vs blank",          "(C $na '#N/A')",   "(C $null)",        False),
    ("#N/A vs zero",           "(C $na '#N/A')",   "(C ([double]0))",  False),
    ("blank vs #N/A",          "(C $null)",        "(C $na '#N/A')",   False),
    ("Int32 4 vs Double 4",    "(C ([int]4))",     "(C ([double]4))",  True),
    ("Double 4 vs Double 5",   "(C ([double]4))",  "(C ([double]5))",  False),
    ("a tiny difference",      "(C ([double]4))",  "(C ([double]4.0000000001))", False),
    ("same driver name",       "(C 'GateB CL-001')", "(C 'GateB CL-001')", True),
    ("different driver names", "(C 'GateB CL-001')", "(C 'GateB CL-002')", False),
    ("case differs",           "(C 'Alpha')",      "(C 'alpha')",      False),
    ("text vs number",         "(C '4')",          "(C ([double]4))",  False),
    ("blank vs blank",         "(C $null)",        "(C '')",           True),
    ("blank vs zero",          "(C $null)",        "(C ([double]0))",  False),
    ("a rho that changed sign", "(C ([double]-0.42))", "(C ([double]0.42))", False),
    ("the same signed rho",    "(C ([double]-0.42))", "(C ([double]-0.42))", True),
]


def _extract_function(name: str) -> str:
    """The function's own source, braces balanced, out of the runner on disk."""
    source = _text()
    start = source.index(f"function {name} {{")
    depth = 0
    for index in range(start, len(source)):
        if source[index] == "{":
            depth += 1
        elif source[index] == "}":
            depth -= 1
            if depth == 0:
                return source[start:index + 1]
    raise AssertionError(f"{name} is not closed")


def _run_comparator(cases: list[tuple[str, str, str, bool]]) -> list[tuple[str, bool]]:
    """Run the runner's OWN comparator over the cases and return what it said."""
    import subprocess
    import tempfile
    body = ["Set-StrictMode -Version 2.0", "$ErrorActionPreference = 'Stop'",
            # THE ONE HELPER IT LEANS ON, copied from the accepted Phase-6 file.
            "function Test-SimBlank {", "    param($Value)",
            "    if ($null -eq $Value) { return $true }",
            "    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }",
            "    return $false", "}"]
    body.append(_extract_function("Test-P83SameCellValue"))
    body.append(_extract_function("Test-P83SameValue"))
    body.append("function C { param($v, [string]$err = '')")
    body.append("  return [pscustomobject]@{ Value = $v; ErrorName = $err; "
                "IsError = ($err -ne '') } }")
    body.append("$na = -2146826246; $ref = -2146826265")
    for label, left, right, _ in cases:
        body.append(f"Write-Output ([string][bool](Test-P83SameCellValue "
                    f"-Left {left} -Right {right}))")
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False,
                                     encoding="utf-8") as handle:
        handle.write("\n".join(body) + "\n")
        path = handle.name
    try:
        done = subprocess.run([PWSH, "-NoProfile", "-File", path],
                              capture_output=True, text=True, timeout=180)
    finally:
        Path(path).unlink(missing_ok=True)
    assert done.returncode == 0, done.stderr[:2000]
    lines = [l.strip() for l in done.stdout.splitlines() if l.strip()]
    assert len(lines) == len(cases), (len(lines), done.stdout[:500])
    return [(cases[i][0], lines[i] == "True") for i in range(len(cases))]


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_80_the_comparator_makes_every_distinction_the_charts_depend_on() -> None:
    """THE SEMANTIC TRUTH TABLE, run against the runner's own comparator."""
    results = _run_comparator(TRUTH_TABLE)
    wrong = [f"{label}: got {got}, expected {case[3]}"
             for (label, got), case in zip(results, TRUTH_TABLE) if got != case[3]]
    assert not wrong, "the comparator does not distinguish:\n  " + "\n  ".join(wrong)


# THE FIVE MOVEMENTS §6 REQUIRES THE PRESERVATION CHECKS TO STILL CATCH once
# the same error code compares equal. The point of a semantic comparison is that
# it stops reporting the non-movement WITHOUT stopping reporting these.
DETECTION_TABLE = [
    # (label, before, after, still-equal?)
    ("a post-window #N/A becomes a fabricated zero",
     "(C $na '#N/A')", "(C ([double]0))", False),
    ("a post-window #N/A becomes a fabricated value",
     "(C $na '#N/A')", "(C ([double]1234.5))", False),
    ("a valid row becomes no-data",
     "(C ([double]1234.5))", "(C $na '#N/A')", False),
    ("a sixth tornado driver appears where there was none",
     "(C $na '#N/A')", "(C 'GateB R-002')", False),
    ("one error code becomes another",
     "(C $na '#N/A')", "(C $ref '#REF!')", False),
    ("a signed rho changes",
     "(C ([double]-0.42))", "(C ([double]-0.31))", False),
    ("a rho keeps its magnitude and loses its sign",
     "(C ([double]-0.42))", "(C ([double]0.42))", False),
    ("a driver category changes",
     "(C 'GateB CL-001')", "(C 'GateB R-002')", False),
    ("a preserved annual value changes",
     "(C ([double]4821017.25))", "(C ([double]4821017.26))", False),
    # AND THE NON-MOVEMENTS THAT MADE RUN 3 RED, which must now be quiet.
    ("an untouched post-window #N/A", "(C $na '#N/A')", "(C $na '#N/A')", True),
    ("an untouched annual value",
     "(C ([double]4821017.25))", "(C ([double]4821017.25))", True),
    ("an untouched driver name", "(C 'GateB CL-001')", "(C 'GateB CL-001')", True),
]


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_80a_the_preservation_checks_still_catch_every_real_movement() -> None:
    """THE HALF THAT MATTERS. Making the same error compare equal must not make
    a fabricated zero, an extra driver, a changed sign or a moved annual value
    compare equal too - and the whole bridge capacity is still compared, so
    nothing is skipped merely for holding #N/A."""
    results = _run_comparator(DETECTION_TABLE)
    wrong = [f"{label}: got equal={got}, expected equal={case[3]}"
             for (label, got), case in zip(results, DETECTION_TABLE) if got != case[3]]
    assert not wrong, "the comparator misses a real movement:\n  " + "\n  ".join(wrong)
    # AND THE PRESERVATION CHECK COMPARES THE WHOLE BLOCK, not a trimmed window.
    body = _function("Compare-P83Bridge")
    assert "for ($index = 0; $index -lt $left.Count; $index++)" in body, (
        "the preservation comparison no longer walks the whole capacity")
    assert "if ($left.Count -ne $right.Count)" in body, (
        "a block that changed length would not be reported")
    for skip in ("IsError) { continue }", "-not $cell.IsError", "#N/A"):
        assert skip not in body, (
            f"the preservation comparison skips rows rather than comparing them: {skip}")


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_81_the_truth_table_is_not_vacuous() -> None:
    """IT MUST CONTAIN BOTH ANSWERS, and the specific pair Run 3 got wrong."""
    assert any(case[3] for case in TRUTH_TABLE)
    assert any(not case[3] for case in TRUTH_TABLE)
    labels = [case[0] for case in TRUTH_TABLE]
    for required in ("same #N/A", "#N/A vs #REF!", "#N/A vs blank", "#N/A vs zero",
                     "Int32 4 vs Double 4", "different driver names"):
        assert required in labels, f"the table does not cover {required}"
    # AND THE OLD RULE WOULD HAVE FAILED IT. `same #N/A` is the pair the runner
    # reported as movement six times over.
    old = dict(zip(labels, [False if "#N/A" in label and "vs" not in label else None
                            for label in labels]))
    assert old["same #N/A"] is False, "the regression this table exists for is gone"


def test_82_one_comparison_owner_and_the_wiring_checks_stay_strict() -> None:
    """NO SEVENTH CALL SITE WITH ITS OWN NOTION OF EQUALITY."""
    code = _code()
    # THE OLD TYPE-GATED HELPERS ARE GONE FROM EVERY CELL-TO-CELL COMPARISON.
    assert "Test-SimSameValue" not in code, (
        "a preservation comparison still uses the type-gated helper")
    body = _function("Invoke-P83TornadoRowChecks")
    assert "Test-SimExactText" not in body, (
        "the tornado still refuses a non-string before comparing")
    assert "Test-P83SameCellValue -Left $bridgeName -Right $sourceName" in body
    assert "Test-P83SameCellValue -Left $bridgeRho -Right $sourceRho" in body
    # PRESERVATION GOES THROUGH THE OWNER.
    assert "Test-P83SameCellValue -Left $left[$index] -Right $right[$index]" in \
        _function("Compare-P83Bridge")
    # AND THE PLOTTED POINT USES THE VALUE HALF OF THE SAME RULE.
    assert "Test-P83SameValue -A $cell.Value -B $point" in _function("Invoke-P83ChartChecks")
    # THE STATE WORDS ARE STILL COMPARED AS TEXT AND MUST STAY THAT WAY: a state
    # cell that stopped being a String is a finding, not a subtype to absorb.
    assert "Test-SimExactText -Actual $live.Value -Expected $Simulation" in \
        _function("Invoke-P83StateChecks")
    # AND THE CHART WIRING IS STILL EXACT STRING WORK ON THE SERIES FORMULA.
    charts = _function("Invoke-P83ChartChecks")
    assert "-cne" in charts, "the range comparison stopped being case-exact"
    assert "$series.Formula" in charts


def test_83_a_non_text_driver_name_is_its_own_finding() -> None:
    """THE ZEROES ARE NOT WAVED AWAY. The contract types driver_name as text, so
    a published label arriving as a number is a statement about the SENSITIVITY
    SHEET, and no comparator change can make it right. It is reported under its
    own name so a reader can tell it from a mirror that disagrees."""
    body = _function("Invoke-P83TornadoRowChecks")
    assert "if ($sourceName.Value -isnot [string]) {" in body
    assert "every published driver name is text" in body
    # AND IT IS A SEPARATE CHECK from the mirror comparison.
    assert body.index("every published driver name is text") > body.index(
        "every plotted driver is the Sensitivity row of the same rank")
    # AND THE OTHER HALF OF THE SAME QUESTION: a ranked driver whose NAME is
    # blank is absent from the chart for a reason that is not "the ranking
    # ended", and the runner says which.
    assert "every ranked driver has a name to plot it under" in body
    contract = yaml.safe_load((SPEC / "sim_contract.yaml").read_text(encoding="utf-8"))
    columns = contract["sim_data"]["sensitivity_records"]["columns"]
    name = next(c for c in columns if c["key"] == "driver_name")
    assert name["value_type"] == "text", (
        "the contract no longer types the driver name as text")


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


def _comparator_ok(code: str) -> None:
    """ONE SEMANTIC EQUALITY, AND IT MAKES EVERY DISTINCTION THE CHARTS NEED."""
    body = re.search(r"function\s+Test-P83SameCellValue\s*\{(.*?)\n\}", code, re.S)
    assert body, "the comparison owner is gone"
    body = body.group(1)
    # ERRORS COMPARED AS ERRORS, BY IDENTITY - not all-equal, not all-unequal.
    assert "if (-not ($Left.IsError -and $Right.IsError)) { return $false }" in body, (
        "an error is compared against a non-error as though it could match")
    assert "[string]$Left.ErrorName -ceq [string]$Right.ErrorName" in body, (
        "two errors are not distinguished by code")
    assert "if ($Left.IsError -or $Right.IsError) { return $false }" not in body, (
        "the comparator refuses an error before reading it")
    value = re.search(r"function\s+Test-P83SameValue\s*\{(.*?)\n\}", code, re.S)
    assert value, "the value half of the rule is gone"
    value = value.group(1)
    # BLANK IS NOT ZERO, AND TEXT IS NOT A NUMBER.
    assert "if ($leftBlank -or $rightBlank) { return ($leftBlank -and $rightBlank) }" in value
    assert "if ($leftText -ne $rightText) { return $false }" in value, (
        "text and numbers are allowed to meet")
    assert "[string]$A -ceq [string]$B" in value, "text is not compared exactly"
    # NO TOLERANCE ANYWHERE IN IT.
    for slack in ("-lt 0.5", "-le 0.000", "[Math]::Round", "-tolerance", "Abs("):
        assert slack not in value, f"a tolerance entered the comparator: {slack}"
    # AND NOTHING ELSE COMPARES TWO CELLS.
    assert "Test-SimSameValue" not in code, "a second notion of equality is back"


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


def _tornado_source_ok(code: str) -> None:
    """THE RANKING IS MIRRORED FROM ITS OWNER, BY KEY, IN PUBLISHED ORDER."""
    assert "-Sensitivity $charts.sensitivity_source" in code, (
        "the tornado does not read the projected sensitivity source")
    for wrong in ("-Sensitivity $simInspection", "-Sensitivity $p8",
                  "-Sensitivity $dash", "-Sensitivity $manifest",
                  "$simInspection.columns"):
        assert wrong not in code, f"the tornado source is read off {wrong}"
    body = re.search(r"function\s+Invoke-P83TornadoRowChecks\s*\{(.*?)\n\}", code, re.S)
    assert body, "the tornado row comparison is gone"
    body = body.group(1)
    assert "$columns['driver_name']" in body and "$columns['rho']" in body, (
        "the tornado no longer looks its columns up by the keys the bridge plots")
    assert "$columns['abs_rho']" not in body, "the tornado reads a magnitude"
    assert "$columns['driver_id']" not in body, "the tornado reads the wrong field"
    for typed in ("'D13'", "'E13'", "'D'", "'E'"):
        assert typed not in body, f"the tornado types a sensitivity address {typed}"
    assert "$sourceRow = $first + $index" in body, (
        "the source row is no longer the positional one")
    for banned in ("Sort-Object", "-Descending", "[array]::Reverse", "$index--"):
        assert banned not in body, f"the tornado reorders for itself: {banned}"
    # A CATEGORY THAT IS NOT TEXT IS A FINDING ABOUT THE SOURCE, and it keeps
    # its own name. Absorbing it into the mirror comparison would let a numeric
    # label pass as soon as both sides carried the same number.
    assert "if ($sourceName.Value -isnot [string]) {" in body, (
        "a non-text published driver name is no longer detected")
    assert "every published driver name is text" in body, (
        "the non-text driver name has stopped being its own check")
    # AND A RANKED DRIVER WITH NO LABEL IS REPORTED RATHER THAN SKIPPED. Without
    # this the chart shows fewer bars than were ranked and says nothing.
    assert "every ranked driver has a name to plot it under" in body, (
        "a published driver with a blank name disappears silently")
    assert "if (-not ((Test-P83Blank -Cell $sourceRho) -or $sourceRho.IsError)) {" in body, (
        "the unnamed-driver finding is not decided by the published rho")


RULES = (_order_ok, _oracle_ok, _recalc_only_ok, _qualification_ok, _fabrication_ok,
         _axis_ok, _freeze_ok, _tornado_source_ok, _comparator_ok)


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
         "    $unnamed = New-Object System.Collections.ArrayList\n"
         "    $plotted = 0",
         "    $unnamed = New-Object System.Collections.ArrayList\n"
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
    # ---- THE SIX FROM THE SECOND WINDOWS RUN ----
    # `.columns` OFF THE WRONG PROJECTION - the crash itself, restored. The
    # Phase-6 gate-B inspection describes the machine sheet and the publication
    # banks; it has no Sensitivity presentation layout and no `columns` at all.
    ("the sensitivity layout is read off the phase-6 machine inspection",
     lambda code: code.replace("-Sensitivity $charts.sensitivity_source",
                               "-Sensitivity $simInspection")),
    # THE RHO POINTED AT THE MAGNITUDE. `abs_rho` sits in the very next column,
    # so every driver would still plot - and every negative one would lose its
    # direction, which is the one thing a tornado exists to show.
    ("the tornado reads the absolute magnitude",
     lambda code: code.replace("$columns['rho']", "$columns['abs_rho']", 1)),
    # THE NAME POINTED AT ANOTHER PUBLISHED FIELD.
    ("the driver name is read from the wrong column",
     lambda code: code.replace("$columns['driver_name']", "$columns['driver_id']", 1)),
    # THE ADDRESS TYPED INTO THE RUNNER.
    ("a sensitivity address is hard-coded",
     lambda code: code.replace(
         "            -Address ($columns['driver_name'] + [string]$sourceRow)",
         "            -Address ('D' + [string]$sourceRow)", 1)),
    # THE TOP-N REORDERED IN POWERSHELL. Phase 7 ranked these; a runner that
    # re-ordered them would agree with itself about a bridge that had drifted.
    ("the top N is reordered in the runner",
     lambda code: code.replace(
         "        $sourceRow = $first + $index",
         "        $sourceRow = $first + ($names.Count - 1 - $index)", 1)),
    # AND THE POSITIONAL WALK REPLACED BY A SORT.
    ("the tornado sorts the published rows",
     lambda code: code.replace(
         "    $unnamed = New-Object System.Collections.ArrayList\n"
         "    $plotted = 0",
         "    $unnamed = New-Object System.Collections.ArrayList\n"
         "    $plotted = 0\n"
         "    $names = @($names | Sort-Object)", 1)),
    # ---- THE FIVE FROM THE THIRD WINDOWS RUN ----
    # THE COMPARATOR BACK AS IT WAS - the rule that called the same #N/A
    # movement six times over.
    ("an error is refused before it is read",
     lambda code: code.replace(
         "    if ($Left.IsError -or $Right.IsError) {\n"
         "        if (-not ($Left.IsError -and $Right.IsError)) { return $false }\n"
         "        return ([string]$Left.ErrorName -ceq [string]$Right.ErrorName)\n"
         "    }",
         "    if ($Left.IsError -or $Right.IsError) { return $false }", 1)),
    # THE OPPOSITE OVERCORRECTION, and the worse one: all errors equal. A
    # post-window #N/A becoming #REF! is a broken reference, and this would call
    # it preserved.
    ("every error is treated as the same error",
     lambda code: code.replace(
         "        return ([string]$Left.ErrorName -ceq [string]$Right.ErrorName)",
         "        return $true", 1)),
    # NO-DATA BECOMING A FABRICATED VALUE, ACCEPTED. An error matched against a
    # number is the fabricated point this layer exists to refuse.
    ("an error is allowed to equal a value",
     lambda code: code.replace(
         "        if (-not ($Left.IsError -and $Right.IsError)) { return $false }",
         "        if (-not ($Left.IsError -and $Right.IsError)) { return $true }", 1)),
    # BLANK ABSORBED INTO ZERO - Excel plots an empty reference as zero, which
    # is the whole reason the bridge emits NA() instead.
    ("blank is allowed to equal zero",
     lambda code: code.replace(
         "    if ($leftBlank -or $rightBlank) { return ($leftBlank -and $rightBlank) }",
         "    if ($leftBlank -and $rightBlank) { return $true }", 1)),
    # TEXT AND NUMBERS ALLOWED TO MEET, which would let a driver named '4' pass
    # for a rho of 4.
    ("text is allowed to equal a number",
     lambda code: code.replace(
         "    if ($leftText -ne $rightText) { return $false }",
         "    if ($leftText -ne $rightText) { return $true }", 1)),
    # A TOLERANCE, which would hide a real chart payload change.
    ("a tolerance is applied to plotted values",
     lambda code: code.replace(
         "    return ([double]$A -eq [double]$B)",
         "    return ([Math]::Abs([double]$A - [double]$B) -lt 0.5)", 1)),
    # AND THE NON-TEXT DRIVER NAME QUIETLY ABSORBED instead of reported.
    ("a non-text driver name stops being reported",
     lambda code: code.replace(
         "        if ($sourceName.Value -isnot [string]) {\n"
         "            $null = $notText.Add('row ' + [string]($index + 1) + ': ' +\n"
         "                                 (Format-P83Cell $sourceName))\n"
         "        }", "", 1)),
    # A RANKED DRIVER WITH NO LABEL, SILENTLY DROPPED - the chart would show
    # fewer bars than were ranked and nothing would say so.
    ("an unnamed ranked driver disappears silently",
     lambda code: code.replace(
         "            if (-not ((Test-P83Blank -Cell $sourceRho) -or $sourceRho.IsError)) {",
         "            if ($false) {", 1)),
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
    """SO THE THIRTY REFUSALS ABOVE ARE REFUSALS OF THE MUTATION, not of the
    fixture."""
    code = _code()
    for rule in RULES:
        rule(code)
    projection = _projection()
    assert len(projection["charts"]) == 4
    assert projection["sensitivity_endpoint"].startswith("PCCM_Run")
