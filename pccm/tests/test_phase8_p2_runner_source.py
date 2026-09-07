#!/usr/bin/env python3
"""P8-2: the minimal Windows runner for the Dashboard executive summary.

WHAT IT IS FOR. P8-2 built a mirror and proved it statically. The claim is that
every analytical Dashboard cell holds `=IF(Results!$D$nn="","",Results!$D$nn)`
and therefore says exactly what Results says, including when Results says
nothing. Whether that holds IN EXCEL, across the state transitions P8-1 was
accepted on, is a runtime question and nothing on Linux can answer it.

WHAT THESE CONTROLS EXIST TO CATCH. A mirroring runner fails in one specific
way: it manufactures the agreement it is supposed to discover. Call an accessor
before reading the worksheet, read Results after reading the Dashboard, or
recompute a total instead of comparing to one, and the run passes while telling
you nothing. So the controls below pin the ORDER of the observation, the ORACLE
(Results, frozen first), and the ABSENCE of any second implementation.

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
RUNNER = WINDOWS / "phase8_p2_dashboard_surface.ps1"
P81 = WINDOWS / "phase8_p1_results_surface.ps1"
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

# THE PARTS REUSED FROM THE ACCEPTED P8-1 RUNNER, renamed and otherwise
# untouched. Reusing means reusing: a paraphrase of a lifecycle that has already
# run on Windows is a new lifecycle.
SHARED_WITH_P81 = (
    "Write-P82Line", "Add-P82Check", "Invoke-P82Release", "Get-P82SourceRevision",
    "Invoke-P82Endpoint", "Set-P82NamedText", "Get-P82Register",
    "Get-P82RegisterColumnIndex", "Get-P82RegisterRowIndex",
)

# THE SIX PART BANNERS, written as literals in the runner precisely so a control
# can find each part without parsing a string concatenation.
BANNERS = (
    "PART 0 - NOTHING HAS RUN",
    "PART A - THE ACCEPTED FIXTURE, THROUGH THE ACCEPTED WORKFLOW",
    "PART B - THE SELECTOR MOVES, RECALCULATION ONLY",
    "PART C - THE ANNUAL ENDPOINT ALONE",
    "PART D - ITERATIONS CHANGE, RECALCULATION ONLY",
    "PART E - AN INVALID MODEL",
)

PARTS = ("part 0", "part A", "part B", "part C", "part D", "part E")

_CACHE: dict = {}


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    return accepted._ps_code(RUNNER)


def _functions() -> dict[str, str]:
    if "fns" not in _CACHE:
        code = _code()
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
        _CACHE["fns"] = out
    return _CACHE["fns"]


def _function(name: str) -> str:
    body = _functions().get(name)
    assert body is not None, f"{name} is not defined in the P8-2 runner"
    return body


def _session() -> str:
    """The executable session, from the first COM acquisition to the verdict."""
    code = _code()
    return code[code.index("$rel = New-ReleaseLedger"):code.index("# VERDICT")]


def _part(name: str) -> str:
    """One part's body, from its banner to the next part's."""
    code = _code()
    starts = []
    for banner in BANNERS:
        marker = f"Write-P82Line '{banner}'"
        assert marker in code, f"the runner has no banner for {banner!r}"
        starts.append(code.index(marker))
    assert starts == sorted(starts), starts
    index = {"part 0": 0, "part A": 1, "part B": 2, "part C": 3,
             "part D": 4, "part E": 5}[name]
    end = starts[index + 1] if index + 1 < len(starts) else len(code)
    return code[starts[index]:end]


def _projection() -> dict:
    if "dash" not in _CACHE:
        _CACHE["dash"] = json.loads(
            (BUILD / "phase8_dashboard_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["dash"]


# ===========================================================================
# A. SCOPE, REUSE AND THE 5.1 FLOOR
# ===========================================================================
def test_01_the_runner_exists_and_is_one_file() -> None:
    assert RUNNER.is_file()
    assert "P8-2" in _text()
    assert _text().startswith("<#"), "the runner has no synopsis"


def test_02_only_the_three_definition_only_files_are_dot_sourced() -> None:
    """A DOT-SOURCE THAT RUNS A SCENARIO would start a second Excel session
    inside this one. These three define functions and run nothing."""
    sourced = re.findall(r"^\.\s+\(Join-Path \$scriptDir '([^']+)'\)", _code(), re.M)
    assert sourced == [path.name for path in DOT_SOURCED], sourced


def test_03_the_ten_helpers_are_byte_identical_to_the_accepted_source() -> None:
    """COPIED, NOT PARAPHRASED - and copied from the same file P8-1 copied from,
    so the tree carries one behaviour rather than three."""
    timing = accepted._ps_code(TIMING)
    mine = _code()
    for name in COPIED_HELPERS:
        pattern = rf"^function\s+{re.escape(name)}\s*\{{.*?^\}}"
        theirs = re.search(pattern, timing, re.S | re.M)
        ours = re.search(pattern, mine, re.S | re.M)
        assert theirs and ours, name
        assert hashlib.sha256(theirs.group(0).encode()).hexdigest() == \
               hashlib.sha256(ours.group(0).encode()).hexdigest(), (
            f"{name} is not byte-identical to the accepted timing harness")


def test_04_the_reused_p81_parts_are_the_p81_ones_renamed() -> None:
    """THE LIFECYCLE HAS ALREADY RUN ON WINDOWS. Reusing it means reusing it;
    the only permitted difference is the P81 -> P82 rename."""
    p81 = accepted._ps_code(P81)
    mine = _code()
    for name in SHARED_WITH_P81:
        original = name.replace("P82", "P81")
        pattern_theirs = rf"^function\s+{re.escape(original)}\s*\{{.*?^\}}"
        pattern_ours = rf"^function\s+{re.escape(name)}\s*\{{.*?^\}}"
        theirs = re.search(pattern_theirs, p81, re.S | re.M)
        ours = re.search(pattern_ours, mine, re.S | re.M)
        assert theirs and ours, name
        assert theirs.group(0).replace("P81", "P82") == ours.group(0), (
            f"{name} diverged from the accepted P8-1 {original}")


def test_05_the_runner_carries_no_powershell_6_construct() -> None:
    """WINDOWS POWERSHELL 5.1 IS THE TARGET SHELL, and the harnesses that have
    actually run there are the evidence for what it accepts."""
    code = _code()
    for construct in ("??", "?.", "&&", "||", "-Parallel", "-LeafBase",
                      "-AsHashtable", "utf8NoBOM", "$IsWindows", "$PSStyle"):
        assert construct not in code, f"the runner uses {construct!r}"
    for line in code.splitlines():
        count = accepted._join_path_positional_count(line)
        assert count is None or count <= 2, f"Join-Path with >2 positionals: {line}"


def test_06_no_comma_expression_is_an_operand_of_arithmetic() -> None:
    """THE DEFECT THAT KILLED P8-1'S SECOND WINDOWS RUN. PowerShell's comma
    binds tighter than its arithmetic, so `@(0, 1, [int]$x - 1)` is an
    Object[] minus an integer. The detector is the accepted one."""
    import test_phase8_p1_runner_source as p81_controls
    found = p81_controls._comma_bound_arithmetic(_text())
    assert not found, f"the P8-2 runner carries comma-bound arithmetic: {found}"


# ===========================================================================
# B. THE ORACLE IS RESULTS - NO SECOND IMPLEMENTATION ANYWHERE
# ===========================================================================
def test_10_every_compared_cell_comes_from_the_projection() -> None:
    """NOT ONE ADDRESS IS SPELLED IN THIS FILE. The projection carries, for each
    mirrored row, the Dashboard cell and the Results cell it must equal; a
    runner that typed either would be a second declaration of something the
    manifest owns, and P7-4 is what that costs when the first one moves."""
    code = _code()
    pairs = _function("Get-P82MirrorPairs")
    assert "$entry.cells.$measure" in pairs
    assert "$cells.dashboard" in pairs and "$cells.results" in pairs
    # NO WORKSHEET ADDRESS LITERAL ANYWHERE IN THE EXECUTABLE BODY.
    for literal in re.finditer(r"['\"]\$?[A-H]\$?(\d{1,3})['\"]", code):
        pytest.fail(f"the runner spells a cell address: {literal.group(0)}")
    for banned in ("Results!D", "Results!$D", "Dashboard!D", "!D4", "!D5"):
        assert banned not in code, f"the runner types the address {banned!r}"


def test_11_the_runner_reads_only_results_and_the_dashboard() -> None:
    """A PRESENTATION ACCEPTANCE THAT REACHED INTO THE MACHINE SHEET would be
    proving the wrong layer's property, and would need a second oracle to do it.
    The sheet name always comes from a projection, never from a literal."""
    code = _code()
    for machine in ("_SimData", "_Calc"):
        # BOUNDARY-AWARE ON PURPOSE. `_Calc` is a substring of
        # PCCM_CalculationStatus, which is an ENDPOINT, not a sheet read, and a
        # naive `in` convicted the compile prerequisite of reading a worksheet.
        hit = re.search(rf"(?<![A-Za-z0-9_]){re.escape(machine)}(?![A-Za-z0-9_])", code)
        assert hit is None, f"the runner names the sheet {machine!r}: {hit.group(0)!r}"
    for reader in ("Get-SimRawCell", "Get-SimBankBlock", "Get-SimSummaryValue",
                   "Get-Phase6State", "Get-SimIterationCell"):
        assert reader not in code, f"the runner reads the machine sheet via {reader}"
    reader = _function("Get-P82Cell")
    assert "param($Workbook, [string]$SheetName, [string]$Address)" in reader, (
        "the reader no longer takes its sheet as an argument")
    # EVERY CALL NAMES ITS SHEET FROM A PROJECTION OR FROM A PAIR.
    for call in re.finditer(r"Get-P82Cell[^\n]*-SheetName\s+(\S+)", code):
        source = call.group(1)
        assert source.startswith("(") or source.startswith("$"), source
        assert "'" not in source and '"' not in source, (
            f"a sheet name is typed rather than projected: {source}")


def test_12_the_runner_recomputes_no_business_quantity() -> None:
    """DASHBOARD ACCEPTANCE TESTS MIRRORING. A runner that subtracted a base
    from a total, summed a profile, took a percentile or compared a difference
    to an allowance would be a second implementation of something Results owns -
    and it would agree with itself about a sheet that had drifted."""
    code = _code()
    for owned in ("quantile_", "contingency_ladder", "Get-P82IdentityAllowance",
                  "identity_absolute_floor", "conditioning", "SimStatsSelected",
                  "ladder", "Test-SimExactDouble -Actual $total"):
        assert owned not in code, f"the runner reaches for {owned!r}"
    # NO ARITHMETIC OVER A COMPARED VALUE. The only arithmetic permitted is over
    # loop indices and row numbers, never over a cell's Value.
    for line in code.splitlines():
        if ".Value" not in line:
            continue
        stripped = line.strip()
        assert not re.search(r"\.Value\s*[-+*/]\s", stripped), (
            f"the runner does arithmetic on a cell value: {stripped}")


def test_13_no_state_word_or_verdict_is_typed_in_the_runner() -> None:
    """THE VOCABULARY IS THE PROJECTION'S. A runner that typed CURRENT would be
    a second authority for a word the contract owns, and it would keep passing
    after the contract renamed it."""
    code = _code()
    for word in ("'CURRENT'", '"CURRENT"', "'HISTORICAL'", '"HISTORICAL"',
                 "'NOT PRODUCED'", '"NOT PRODUCED"', "'OTHER Px'", '"OTHER Px"',
                 "'Reconciled'", '"Reconciled"', "'NOT RECONCILED'", "'INVALID'",
                 '"INVALID"', "'STALE'", '"STALE"'):
        assert word not in code, f"the runner types the state word {word}"
    for projected in ("$p7.handoff.distribution_states", "$p7.handoff.profile_states",
                      "$p7.handoff.inconsistent_stamp_state",
                      "$p7.model_states.derived_status"):
        assert projected in code, f"the runner does not project {projected}"


# ===========================================================================
# C. THE OBSERVATION ORDER
# ===========================================================================
def test_20_the_observation_orchestrator_runs_the_six_steps_in_order() -> None:
    """THE ORDER IS THE MEASUREMENT. Recalculate, freeze Results, freeze the
    Dashboard, compare - with nothing between them. A Results value re-read
    AFTER the Dashboard would be a value the Dashboard could have changed."""
    body = _function("Invoke-P82Observation")
    steps = ["Invoke-P82Recalculate", "Get-P82ResultsFrozen",
             "Get-P82DashboardFrozen", "Invoke-P82MirrorChecks"]
    positions = [body.index(step) for step in steps]
    assert positions == sorted(positions), list(zip(steps, positions))
    # NOTHING OUT OF CELL INSIDE THE OBSERVATION.
    for forbidden in ("$Excel.Run", "Invoke-P82Endpoint", "Set-NamedValue",
                      "Set-P82NamedText", "Set-TableCell", "Set-Phase5Fixture"):
        assert forbidden not in body, (
            f"{forbidden} happens inside the frozen observation")
    # AND NO SECOND RECALCULATION between the two freezes.
    between = body[body.index("Get-P82ResultsFrozen"):body.index("Invoke-P82MirrorChecks")]
    assert "Calculate" not in between, "the workbook recalculates between the freezes"


def test_21_results_is_frozen_before_the_dashboard_in_every_part() -> None:
    """PROVED PER PART, not only in the orchestrator, so a part that froze by
    hand in the wrong order is caught too."""
    for name in PARTS:
        body = _part(name)
        assert "Invoke-P82Observation" in body, f"{name} makes no observation"
        for hand_rolled in ("Get-P82DashboardFrozen", "Get-P82ResultsFrozen"):
            assert hand_rolled not in body, (
                f"{name} freezes by hand instead of through the orchestrator")


def test_22_the_state_checks_read_the_frozen_capture_not_the_sheet() -> None:
    """AN ASSERTION THAT RE-READ THE SHEET would be reading it after the
    Dashboard had been read, which is a different moment and a different fact."""
    for name in ("Invoke-P82StateChecks", "Invoke-P82SelectorChecks",
                 "Invoke-P82HeadlineChecks", "Invoke-P82QualifierChecks",
                 "Test-P82NoInventedStatus"):
        body = _function(name)
        assert "Get-P82Cell" not in body, f"{name} re-reads the sheet"
        assert ("Get-P82Frozen" in body) or ("$Observation." in body), name


@pytest.mark.parametrize("part", ["part B", "part D"])
def test_23_no_endpoint_runs_before_the_observation_in_the_recalc_only_parts(
        part: str) -> None:
    """THE TWO PARTS WHOSE WHOLE QUESTION IS WHAT A RECALCULATION ALONE DID. An
    endpoint here would republish the very thing the part exists to find
    unchanged, and the run would pass having proved nothing."""
    body = _part(part)
    before = body[:body.index("Invoke-P82Observation")]
    for endpoint in ("Invoke-P82Endpoint", "Invoke-Phase6Simulation",
                     "Invoke-Phase5ProductionOperation", "$excel.Run("):
        assert endpoint not in before, (
            f"{part} invokes {endpoint} before the observation")
    # AND NOT ANYWHERE ELSE IN THE PART EITHER, except the one diagnostic part D
    # takes AFTER its observation.
    after = body[body.index("Invoke-P82Observation"):]
    assert "Invoke-P82Endpoint" not in after, f"{part} invokes an endpoint at all"
    if part == "part B":
        assert "$excel.Run(" not in after, "part B calls out of cell"
    else:
        assert after.count("$excel.Run(") == 1, (
            "part D takes more than the one post-observation diagnostic")
        assert after.index("$excel.Run(") > after.index("Compare-P82Witness"), (
            "part D's diagnostic runs before its witness comparison")


def test_24_the_transition_always_precedes_the_observation() -> None:
    """STEP 1 BEFORE STEP 2. A part that observed and then transitioned would be
    reporting the previous part's state under this part's name."""
    transitions = {
        "part A": "Set-Phase5Fixture",
        "part B": "Set-P82NamedText",
        "part C": "Invoke-P82Endpoint",
        "part D": "Set-NamedValue -Workbook $wb -DefinedName $iterationsControl",
        "part E": "Set-TableCell",
    }
    for name, transition in transitions.items():
        body = _part(name)
        assert transition in body, f"{name} performs no transition"
        assert body.index(transition) < body.index("Invoke-P82Observation"), (
            f"{name} observes before it transitions")


# ===========================================================================
# D. THE DISTINCTIONS, THE PARTS AND THE RESERVED REGION
# ===========================================================================
def test_30_the_six_parts_are_present_and_in_order() -> None:
    code = _code()
    positions = [code.index(f"Write-P82Line '{banner}'") for banner in BANNERS]
    assert positions == sorted(positions), positions
    assert len(set(positions)) == 6


def test_31_the_selected_level_and_the_profile_px_are_compared_separately() -> None:
    """TWO DIFFERENT QUESTIONS, and part B is the part where they differ. A
    runner that read either from the other's row would make that part pass
    without observing anything."""
    body = _function("Invoke-P82SelectorChecks")
    assert "-Key 'selected_confidence_level'" in body
    assert "-Key 'profile_px'" in body
    assert "$expectSame" in body, (
        "the two fields are not compared as separately expected values")
    part_b = _part("part B")
    assert "-SelectedLabel $secondLabel -ProfilePx $firstLabel" in part_b, (
        "part B does not expect the selector to move while the profile stays")


def test_32_the_total_and_the_contingency_are_asserted_distinct() -> None:
    """W5 READ ONE AS THE OTHER. Nothing on the Dashboard recomputes the
    difference, so the distinction is asserted rather than inferred."""
    body = _function("Invoke-P82HeadlineChecks")
    assert "-Key 'selected_total'" in body and "-Key 'contingency'" in body
    assert "is not the contingency wearing its name" in body
    assert "foreach ($measure in @('nominal', 'pv'))" in body, (
        "the headline is only checked on one measure")


def test_33_the_four_state_lines_and_the_year_count_are_all_asserted() -> None:
    body = _function("Invoke-P82StateChecks")
    for key in ("distribution_state", "profile_state", "profile_px", "year_count"):
        assert f"-Key '{key}'" in body, key
    assert "the two state lines are not the same word" in body, (
        "nothing stops the two state answers being collapsed into one")


def test_34_the_reconciliation_qualifier_is_required_to_arrive_whole() -> None:
    body = _function("Invoke-P82QualifierChecks")
    assert "qualifier_suffix" in body
    assert "$text.Contains($qualifier)" in body
    assert "$text.Contains($ProfileState)" in body, (
        "the qualifier is not required to name the state it qualifies")
    # AND IT IS ASKED IN EVERY PART THAT HAS A PUBLISHED RESULT.
    for name in ("part A", "part B", "part C", "part D", "part E"):
        assert "Invoke-P82QualifierChecks" in _part(name), name


def test_35_the_chart_region_is_proved_empty_live() -> None:
    """THREE SEPARATE ABSENCES, because they fail for different reasons: a chart
    object is P8-3 started early, a shape is one pasted, and a value in the
    reserved rows is the summary having grown into space the next step needs."""
    body = _function("Test-P82ChartRegion")
    assert "ChartObjects()" in body and "$sheet.Shapes" in body
    assert "$region.first_row" in body and "$region.last_row" in body
    assert "no chart object" in body and "no picture or shape" in body
    # PROBED AT THE START AND AT THE END, so a part that drew something is caught.
    assert "Test-P82ChartRegion" in _part("part 0")
    assert "Test-P82ChartRegion" in _part("part E")


def test_36_part_0_requires_blanks_rather_than_fabricated_zeros() -> None:
    body = _part("part 0")
    assert "blank, not a fabricated zero" in body
    assert "Test-P82Blank" in body
    assert "-ExpectedPx $null" in body, "part 0 does not expect a blank profile Px"
    assert "-ExpectedYears 0" in body


def test_37_part_e_asserts_the_absence_of_an_invented_status() -> None:
    """P8-2 REPORTED THAT RESULTS PUBLISHES NO LIVE DETERMINISTIC STATUS. The
    runner must assert that absence rather than quietly start expecting a row
    nobody built - and must not conjure the word INVALID onto the Dashboard."""
    body = _part("part E")
    assert "Test-P82NoInventedStatus" in body
    assert "NO INVALID INDICATOR IS EXPECTED" in _text(), (
        "the runner does not say why it asserts an absence here")
    sweep = _function("Test-P82NoInventedStatus")
    assert "$mirror.Value" in sweep, (
        "the sweep does not compare what the Dashboard says against what Results said")
    # THE ONLY CALCULATION-STATUS ASSERTION IS OUT OF CELL AND AFTER THE FREEZE.
    assert body.index("Invoke-P82Observation") < body.index("PCCM_CalculationStatus")


def test_38_the_payload_is_witnessed_through_results_not_around_it() -> None:
    """PARTS B AND D CHANGE NOTHING AND MUST BE SEEN TO CHANGE NOTHING - through
    the surface that presents the payload, because reaching into the machine
    sheet would need a second oracle."""
    witness = _function("Get-P82ResultsWitness")
    assert "$P8.run_stamp.fields" in witness and "$P8.summary.metrics" in witness
    assert "$P8.annual.first_row" in witness
    assert "_SimData" not in witness
    for name in ("part B", "part D", "part E"):
        assert "Compare-P82Witness" in _part(name), name
    # AND THE WITNESS IS THE ONLY PAYLOAD ORACLE. The P8-1 run-invariant helper
    # was deliberately NOT carried over: it reads the persisted block through
    # Get-Phase6State, which would have made this runner read three sheets to
    # prove a property the presenting surface already shows.
    assert "Get-P82RunInvariants" not in _code()
    assert "Add-P82InvariantChecks" not in _code()


# ===========================================================================
# E. MUTATIONS - each is a way this runner could pass while proving nothing
# ===========================================================================
def _order_ok(code: str) -> None:
    body = re.search(r"function\s+Invoke-P82Observation\s*\{(.*?)\n\}", code, re.S)
    assert body, "the orchestrator is gone"
    text = body.group(1)
    positions = [text.index(step) for step in
                 ("Invoke-P82Recalculate", "Get-P82ResultsFrozen",
                  "Get-P82DashboardFrozen", "Invoke-P82MirrorChecks")]
    assert positions == sorted(positions), "the observation order is wrong"
    for forbidden in ("$Excel.Run", "Invoke-P82Endpoint"):
        assert forbidden not in text, f"{forbidden} is inside the observation"


def _oracle_ok(code: str) -> None:
    for machine in ("_SimData", "_Calc"):
        hit = re.search(rf"(?<![A-Za-z0-9_]){re.escape(machine)}(?![A-Za-z0-9_])", code)
        assert hit is None, f"the runner names the sheet {machine}"
    for reader in ("Get-SimRawCell", "Get-SimBankBlock", "Get-Phase6State"):
        assert reader not in code, f"the runner reads the machine sheet via {reader}"


def _recalc_only_ok(code: str) -> None:
    for banner in ("PART B - ", "PART D - "):
        start = code.index(f"Write-P82Line '{banner}")
        body = code[start:code.index("Invoke-P82Observation", start)]
        for endpoint in ("Invoke-P82Endpoint", "Invoke-Phase6Simulation", "$excel.Run("):
            assert endpoint not in body, f"an endpoint runs before {banner}observation"


def _fields_stay_apart(code: str) -> None:
    body = re.search(r"function\s+Invoke-P82SelectorChecks\s*\{(.*?)\n\}", code, re.S)
    assert body, "the selector check is gone"
    assert "-Key 'selected_confidence_level'" in body.group(1)
    assert "-Key 'profile_px'" in body.group(1)
    head = re.search(r"function\s+Invoke-P82HeadlineChecks\s*\{(.*?)\n\}", code, re.S)
    assert head, "the headline check is gone"
    assert "is not the contingency wearing its name" in head.group(1)


RULES = (_order_ok, _oracle_ok, _recalc_only_ok, _fields_stay_apart)


@pytest.mark.parametrize("name,mutate", [
    # THE ACCESSOR CALLED BEFORE THE WORKSHEET IS READ - P8-1's disclosed defect,
    # reintroduced. It manufactures the agreement the observation should find.
    ("an out-of-cell call moves inside the observation",
     lambda code: code.replace(
         "    $results = Get-P82ResultsFrozen -Workbook $Workbook -Pairs $Pairs",
         "    $null = $Excel.Run('PCCM_CalculationStatus')\n"
         "    $results = Get-P82ResultsFrozen -Workbook $Workbook -Pairs $Pairs", 1)),
    # THE DASHBOARD FROZEN FIRST - Results then re-read afterwards, which is a
    # different moment and a value the Dashboard could have influenced.
    ("the dashboard is frozen before results",
     lambda code: code.replace(
         "    $results = Get-P82ResultsFrozen -Workbook $Workbook -Pairs $Pairs\n",
         "", 1).replace(
         "    $verdict = Invoke-P82MirrorChecks",
         "    $results = Get-P82ResultsFrozen -Workbook $Workbook -Pairs $Pairs\n"
         "    $verdict = Invoke-P82MirrorChecks", 1)),
    # A SECOND ORACLE - the machine sheet read directly, so the runner no longer
    # proves the presentation layer.
    ("the runner reads the machine sheet",
     lambda code: code.replace(
         "function Get-P82ResultsWitness {",
         "function Get-P82ResultsWitness {\n    # _SimData\n    $null = Get-SimRawCell", 1)),
    # AN ENDPOINT IN A RECALCULATION-ONLY PART - the part passes having
    # republished the thing it was supposed to find unchanged.
    ("part B runs an endpoint before observing",
     lambda code: code.replace(
         "    $observationB = Invoke-P82Observation",
         "    $null = Invoke-P82Endpoint -Excel $excel -Endpoint 'PCCM_Calculate'\n"
         "    $observationB = Invoke-P82Observation", 1)),
    ("part D runs an endpoint before observing",
     lambda code: code.replace(
         "    $observationD = Invoke-P82Observation",
         "    $null = Invoke-P82Endpoint -Excel $excel -Endpoint 'PCCM_Calculate'\n"
         "    $observationD = Invoke-P82Observation", 1)),
    # THE TWO Px QUESTIONS COLLAPSED INTO ONE - part B stops being observable.
    ("the selected level is read from the profile Px row",
     lambda code: code.replace(
         "    $selectedCell = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'selected_confidence_level'",
         "    $selectedCell = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'profile_px'", 1)),
    # THE TWO LADDERS COLLAPSED - the W5 defect, and nothing here would notice.
    ("the total and the contingency stop being compared",
     lambda code: code.replace("is not the contingency wearing its name",
                               "is a number", 1)),
])
def test_50_each_way_of_passing_while_proving_nothing_is_refused(
        name: str, mutate) -> None:
    """EACH MUTATION APPLIED TO A COPY OF THE SOURCE IN MEMORY, with the rules
    above re-run over it. Nothing on disk changes. A mutation that survives every
    rule means the rules are decoration."""
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
    """SO THE SEVEN REFUSALS ABOVE ARE REFUSALS OF THE MUTATION, not of the
    fixture."""
    code = _code()
    for rule in RULES:
        rule(code)
    # AND THE PROJECTION THE RUNNER READS EXISTS AND CARRIES PAIRS.
    projection = _projection()
    assert projection["sections"], "the Dashboard projection is empty"
    assert any(entry["cells"] for section in projection["sections"]
               for entry in section["rows"])
