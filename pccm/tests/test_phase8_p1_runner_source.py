#!/usr/bin/env python3
"""P8-1: the minimal Windows runner for the Results output surface.

WHAT IT IS FOR. P8-1 built a sheet and corrected where its state comes from.
Both are source-proved and neither is runtime-proved, and one question has to be
answered before the rest is worth reading: two of the four handoff accessors
reach a status derivation that PERSISTS two rows, and Excel does not let a
function called from a cell change the workbook. So the runner's first live
assertion, before a single simulation, is that the four state cells evaluate.

THE GATE IS A STOP, NOT A CHECK. If those cells do not evaluate, the run ends
there and returns the evidence. A runner that carried on - or that read the
formatted text instead, or retried through the endpoints - would destroy exactly
the observation the round was authorised to make. Controls below pin the stop.

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
sys.path.insert(0, str(PCCM_ROOT / "builder"))

import pytest  # noqa: E402

import test_phase7_acceptance_harness_source as accepted  # noqa: E402

WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
RUNNER = WINDOWS / "phase8_p1_results_surface.ps1"
W8 = WINDOWS / "phase7_w8_refusal.ps1"
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
CALLED_DIRECTLY = ("Get-NamedValue", "Set-NamedValue", "Get-TableBody", "Set-TableCell")

# THE PARTS THE LIFECYCLE AND THE READERS ARE REUSED FROM. W8 is closed Windows
# evidence; reusing it means reusing it, not paraphrasing it.
SHARED_WITH_W8 = (
    "Write-P81Line", "Add-P81Check", "Invoke-P81Release", "Get-P81SourceRevision",
    "Get-P81RowsInGroup", "Get-P81RunInvariants", "Add-P81InvariantChecks",
    "ConvertTo-P81ColumnNumber", "ConvertFrom-P81ColumnNumber", "Get-P81AnnualStamp",
    "Get-P81AnnualRecord", "Get-P81AnnualSurface", "Get-P81BankCapture",
    "Compare-P81Surface", "Get-P81IterationBlock", "Compare-P81IterationGrid",
    "Get-P81Handoff", "Invoke-P81Endpoint", "Set-P81NamedText",
    "Get-P81IdentityAllowance", "Get-P81Register", "Get-P81RegisterColumnIndex",
    "Get-P81RegisterRowIndex",
)

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
    assert body is not None, f"{name} is not defined in the P8-1 runner"
    return body


def _own_code() -> str:
    code = _code()
    for name in COPIED_HELPERS:
        code = code.replace(_function(name), "")
    return code


def _projection() -> dict:
    if "p8" not in _CACHE:
        _CACHE["p8"] = json.loads(
            (BUILD / "phase8_results_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["p8"]


# ===========================================================================
# THE COMMA-BINDS-TIGHTER-THAN-ARITHMETIC DETECTOR
# ===========================================================================
# THE DEFECT IT EXISTS FOR, IN ONE LINE. PowerShell's comma binds TIGHTER than
# its arithmetic, so `@(0, 1, [int]$x - 1)` is `(0, 1, [int]$x) - 1` - an
# Object[] minus an integer. The first Windows run of P8-1 died on precisely
# that, one statement past the gate it had just passed:
#
#     Method invocation failed because [System.Object[]] does not contain a
#     method named 'op_Subtraction'.
#
# The cast was present and correct; it simply applied to an operand that was
# never the left-hand side. Nothing in a Python test can parse PowerShell
# properly, so this reads the shape rather than the grammar: inside a bracket
# group where a comma IS the array operator, a top-level arithmetic operator is
# the trap. A METHOD CALL's argument list is not a comma expression - the
# arguments are parsed separately - so `[Math]::Min($a + 1, $b)` is safe, and
# the opening bracket is classified by what precedes it.
_ARITHMETIC = set("+-*/%")


def _comma_bound_arithmetic(code: str) -> list[tuple[int, str]]:
    out: list[tuple[int, str]] = []
    stack: list[dict] = []
    index, size, line = 0, len(code), 1
    while index < size:
        char = code[index]
        if char == "\n":
            line += 1
            index += 1
            continue
        if char == "#":
            while index < size and code[index] != "\n":
                index += 1
            continue
        if char in "'\"":
            quote = char
            index += 1
            while index < size:
                if code[index] == "\n":
                    line += 1
                if code[index] == quote:
                    if index + 1 < size and code[index + 1] == quote:
                        index += 2
                        continue
                    break
                if quote == '"' and code[index] == "`":
                    index += 1
                index += 1
            index += 1
            continue
        if char in "([{":
            kind = "group"
            if char == "(":
                # A CALL'S BRACKET TOUCHES ITS NAME. `Foo(` and `[Math]::Min(`
                # are calls; `-Offset (` and `in @(` and a bare `(` are not, and
                # the whitespace is what says so - which is the distinction the
                # first draft of this detector got wrong, and it let a mutation
                # through in exactly that shape.
                adjacent = code[:index]
                if adjacent.rstrip().endswith("@") and adjacent.endswith("@"):
                    kind = "array"
                elif re.search(r"[\w\]\.]$", adjacent) and not adjacent.endswith("$"):
                    kind = "call"
                else:
                    kind = "array"
            elif char == "[":
                before = code[:index]
                kind = "array" if re.search(r"[\w\)\]]$", before) else "type"
            stack.append({"kind": kind, "line": line, "comma": False,
                          "arith": False, "start": index})
            index += 1
            continue
        if char in ")]}":
            if stack:
                top = stack.pop()
                if top["kind"] == "array" and top["comma"] and top["arith"]:
                    out.append((top["line"],
                                code[top["start"]:index + 1].replace("\n", " ")[:90]))
            index += 1
            continue
        if stack and stack[-1]["kind"] in ("array", "call"):
            if char == ",":
                stack[-1]["comma"] = True
            elif char in _ARITHMETIC and stack[-1]["kind"] == "array":
                previous = code[index - 1] if index else " "
                following = code[index + 1] if index + 1 < size else " "
                # A binary operator has whitespace on both sides; `-eq`, a
                # negative literal and a `-Parameter` name do not.
                if previous == " " and following == " ":
                    stack[-1]["arith"] = True
        index += 1
    return out

# ===========================================================================
# A. SCOPE
# ===========================================================================

def test_01_the_runner_exists_and_is_dedicated_to_p8_1() -> None:
    assert RUNNER.exists()
    lines = _text().splitlines()
    assert len(lines) < 2000, f"{len(lines)} lines; P8-1 owns one output surface"
    head = "\n".join(lines[:80])
    assert "P8-1" in head and "Results" in head


def test_02_it_dot_sources_only_definition_only_files() -> None:
    sourced = re.findall(r"^\.\s+\(Join-Path \$scriptDir '([^']+)'\)", _code(), re.M)
    assert sourced == ["com_lifecycle.ps1", "phase5_gate_b_scenarios.ps1",
                       "phase6_gate_b_scenarios.ps1"], sourced
    code = _code()
    assert re.search(r"^\$bootstrap = Join-Path \$scriptDir 'build_stage_b\.ps1'", code, re.M)
    assert re.search(r"^& \$bootstrap -BuildDir \$tempRoot -Force", code, re.M)


def test_03_no_phase_7_scenario_is_re_tested_and_no_later_phase_is_begun() -> None:
    own = _own_code()
    for closed in ("phase7_w1_smoke", "phase7_w2_many_drivers", "phase7_w3_long_years",
                   "phase7_w4_base_simulation", "phase7_w5_annual_success",
                   "phase7_w6_selector_move", "phase7_w7_bank_cycle", "phase7_w8_refusal",
                   "phase7_acceptance_scenarios"):
        assert closed not in own, f"P8-1 reaches the closed {closed}"
    for later in ("Dashboard", "ChartObjects", "AddChart", "SeriesCollection",
                  "PCCM_RunSensitivity", "Histogram", "Tornado"):
        assert later not in own, f"P8-1 reaches {later}, which is a later step"


def test_04_the_accepted_w8_runner_is_untouched() -> None:
    """W8 IS CLOSED WINDOWS EVIDENCE and P8-1 reuses its shape by copying, not by
    editing it. The digest is the one W8's own suite pins."""
    digest = hashlib.sha256(W8.read_bytes()).hexdigest()
    assert digest == "cf4adee6061af83760e6a441f70ae0737456e5ac04775ebce591e9ff57b51306", (
        "the accepted W8 runner has been modified; it is closed evidence")


def test_05_the_shape_shared_with_w8_is_identical_to_w8s() -> None:
    ours = _text()
    theirs = W8.read_text(encoding="utf-8")
    for name in SHARED_WITH_W8:
        mine = re.search(rf"^function {name} \{{(.*?)^\}}", ours, re.M | re.S)
        assert mine, f"{name} is not defined in the P8-1 runner"
        w8name = name.replace("P81", "W8")
        yours = re.search(rf"^function {w8name} \{{(.*?)^\}}", theirs, re.M | re.S)
        assert yours, f"{w8name} is no longer in the accepted W8 runner"
        assert mine.group(1).replace("P81", "W8") == yours.group(1), (
            f"{name} has drifted from the accepted {w8name}; the lifecycle and the "
            "readers were to be reused unchanged")


# ===========================================================================
# B. CLOSURE AND SHELL
# ===========================================================================

def test_10_every_custom_command_the_runner_reaches_is_defined() -> None:
    defined: dict[str, str] = {}
    for path in (RUNNER,) + DOT_SOURCED:
        for name, body in accepted._ps_functions(path).items():
            defined.setdefault(name, body)
    seen: set[str] = set()
    missing: list[tuple[str, str]] = []
    work = [("<runner top level>", accepted._ps_top_level(RUNNER))]
    work += list(accepted._ps_functions(RUNNER).items())
    while work:
        where, body = work.pop()
        for call in sorted(set(accepted._VERB_NOUN.findall(body))):
            if call in accepted.POWERSHELL_BUILTINS:
                continue
            if call not in defined:
                missing.append((where, call))
                continue
            if call in seen:
                continue
            seen.add(call)
            work.append((call, defined[call]))
    assert not missing, "\n  ".join(f"{c} (from {w})" for w, c in sorted(set(missing)))
    assert len(seen) > 40


def test_11_the_copied_helpers_are_verbatim_and_honestly_labelled() -> None:
    timing = TIMING.read_text(encoding="utf-8")
    raw = _text()
    own = _own_code()
    body = "\n".join(l for l in own.splitlines() if not l.strip().startswith("function "))
    for name in COPIED_HELPERS:
        theirs = re.search(rf"^function {name} \{{(.*?)^\}}", timing, re.M | re.S)
        ours = re.search(rf"^function {name} \{{(.*?)^\}}", raw, re.M | re.S)
        assert theirs and ours, name
        assert ours.group(1) == theirs.group(1), f"{name} drifted from the accepted copy"
        called = re.search(rf"(?<![\w\-.$]){name}(?![\w\-])", body)
        if name in CALLED_DIRECTLY:
            assert called, f"{name} is listed as called directly but nothing calls it"
        else:
            assert not called, f"{name} is called after all; the justification is stale"


def test_12_no_powershell_6_or_7_only_construct_is_used() -> None:
    code = _code()
    for label, pattern in accepted.PS51_ONLY_CONSTRUCTS:
        assert not re.search(pattern, code), f"{label} is not available in Windows PowerShell 5.1"
    assert accepted._join_path_positional_count(code) <= 2
    assert "WINDOWS POWERSHELL 5.1" in _text()


def test_13_the_runner_carries_no_dead_function() -> None:
    """A MINIMAL RUNNER HAS NOTHING SPARE IN IT. The ten copied helpers are the
    stated exception - the accepted fixture calls them and its own file does not
    define them - and everything else must be reached."""
    text = _text()
    for name in _functions():
        if name in COPIED_HELPERS:
            continue
        assert text.count(name) >= 2, f"{name} is defined and never called"


# ===========================================================================
# C. THE GATE THAT COMES BEFORE EVERYTHING
# ===========================================================================

def test_20_the_state_evaluation_gate_is_the_first_live_assertion() -> None:
    code = _code()
    gate = code.index("$observation0 = Invoke-P81Observation")
    for later in ("Set-Phase5Fixture", "Invoke-Phase6Simulation",
                  "$p7.command_surface.annual_endpoint", "PCCM_Calculate"):
        assert code.index(later) > gate, (
            f"{later} is reached before the state cells are proved to evaluate")


def test_21_a_failed_gate_stops_the_run_and_is_not_worked_around() -> None:
    code = _code()
    assert "if (-not $observation0.Evaluated) {" in code
    stop = code[code.index("if (-not $observation0.Evaluated) {"):]
    end = stop.index("throw 'the Results state cells did not evaluate")
    assert "$stoppedOnStateEvaluation = $true" in stop[:end], (
        "a failed gate must be recorded as a stop")
    assert stop[end:].startswith("throw"), (
        "a failed gate must end the session, not be logged and passed")
    stop = stop[:end]
    for workaround in ("Resume Next", "-ErrorAction SilentlyContinue", "Text -eq",
                       "PCCM_RunAnnualStochastic", "PCCM_RunSimulation",
                       "Invoke-P81Recalculate", "Get-P81StateCells"):
        assert workaround not in stop, f"the gate is worked around with {workaround}"
    # THE ONE THING THE STOP MAY DO is ask the same accessors outside a cell,
    # after the failed observation is already frozen - and it must be labelled
    # as a diagnostic rather than counted as a check.
    assert "DIAGNOSTIC (not a check)" in stop
    assert "Add-P81Check" not in stop, "the stop block adds a check after the gate failed"
    assert "P8-1: STOPPED AT THE STATE-EVALUATION GATE" in code


def test_22_an_excel_error_is_detected_by_code_and_by_text() -> None:
    """NEITHER ROUTE ALONE IS GUARANTEED. An error cell's Value2 arrives as an
    Int32 in the CVErr band; its Text is the error name. The runner exists to
    settle a question about exactly this, so it reads both."""
    body = _function("Get-P81Cell")
    assert "$script:P81ErrorCodes.ContainsKey($code)" in body
    assert "-ceq [string]$name" in body
    codes = _code()
    for code in ("-2146826273", "-2146826265", "-2146826259", "-2146826246"):
        assert code in codes, f"the CVErr code {code} is not recognised"
    for name in ("'#VALUE!'", "'#REF!'", "'#NAME?'", "'#N/A'"):
        assert name in codes, f"the error name {name} is not recognised"


def test_23_the_text_is_never_the_numeric_authority() -> None:
    """`#,##0` turns 4.5E-13 into `0`. A check that compared the displayed
    string would call a mismatch agreement, which is the one way a reconciliation
    can lie without anything being wrong with it."""
    same = _function("Test-P81SameNumber")
    assert "Text" not in same, "the numeric comparison can see the formatted text"
    assert "$Cell.Value -isnot [double]" in same, (
        "a value that is not a Double must not be coerced into agreement")
    assert "return ([double]$Cell.Value -eq [double]$Expected)" in same, (
        "the comparison is between the two stored Doubles and nothing else")
    for name in ("Invoke-P81ReconciliationChecks", "Invoke-P81AnnualChecks"):
        body = _function(name)
        assert ".Text" not in body, f"{name} compares the formatted text"


# ===========================================================================
# D. THE SURFACE IS READ THROUGH ITS PROJECTION
# ===========================================================================

def test_30_no_results_address_is_written_in_the_runner() -> None:
    """THE P7-4 LESSON. A hand-written address is a second declaration of
    something a manifest already owns, and it goes stale in silence."""
    own = _own_code()
    assert not re.search(r"'[A-H]\$?\d{1,3}'", own), "an absolute Results address is typed"
    for procedure in ("PCCM_ResultsAnnualDistributionState", "PCCM_ResultsAnnualProfileState",
                      "PCCM_ResultsAnnualProfilePx", "PCCM_ResultsAnnualYearCount"):
        assert procedure not in own, f"{procedure} is typed rather than projected"
    assert "$Inspection.state.$key.procedure" in own
    assert "$Inspection.annual.columns" in own
    assert "$Inspection.reconciliation.rows" in own


def test_31_the_projection_is_generated_and_validated() -> None:
    from pccm_builder import load_sim_contract
    from pccm_builder.calc_loader import load_calc_contract
    from pccm_builder.spec_loader import load_spec
    from pccm_builder.structure_loader import load_structure_contract
    from pccm_builder import phase8_results as emitter

    spec = load_spec(SPEC / "workbook.yaml")
    calc = load_calc_contract(SPEC / "calc_contract.yaml")
    sim = load_sim_contract(SPEC / "sim_contract.yaml")
    limits = load_structure_contract(SPEC / "structure_contract.yaml").limits
    rebuilt = emitter.build_phase8_inspection(
        spec, calc, sim.raw, limits.max_generated_year_columns)
    assert rebuilt == _projection(), (
        "phase8_results_inspection.json is not what its generator produces")

    merged = json.loads(json.dumps(rebuilt))
    merged["selected"]["contingency_row"] = merged["selected"]["total_row"]
    with pytest.raises(ValueError, match="different quantities"):
        emitter.validate_phase8_inspection(merged)
    unwrapped = json.loads(json.dumps(rebuilt))
    unwrapped["state"]["profile_state"]["procedure"] = (
        unwrapped["state"]["profile_state"]["accessor"])
    with pytest.raises(ValueError, match="answer once and never again"):
        emitter.validate_phase8_inspection(unwrapped)
    overrun = json.loads(json.dumps(rebuilt))
    overrun["annual"]["row_window"] = overrun["reconciliation"]["heading_row"]
    with pytest.raises(ValueError, match="past the reconciliation heading"):
        emitter.validate_phase8_inspection(overrun)


def test_32_the_identity_rule_comes_from_the_projection() -> None:
    own = _own_code()
    assert "$Inspection.identity" in own
    for invented in ("1e-6", "1e-12", "0.000001", "0.000000000001"):
        assert invented not in own, f"{invented} is a tolerance this runner invented"
    body = _function("Invoke-P81ReconciliationChecks")
    assert "Get-P81IdentityAllowance -Provenance $Inspection.identity" in body


def test_33_the_total_is_proved_not_to_be_the_contingency() -> None:
    """THE W5 SEMANTIC, CHECKED AT RUNTIME. Both are money at the same Px and
    the contingency is smaller by exactly the deterministic base, so a runner
    that read the wrong one would reconcile against a number that looked
    entirely reasonable."""
    body = _function("Invoke-P81ReconciliationChecks")
    # THE ASSIGNMENTS, NOT THE MENTIONS. A function that calls the right reader
    # and then compares something else would satisfy a substring check.
    assert "$total = Get-SimSummaryValue -Workbook $Workbook" in body, (
        "the authoritative total is not read from the summary block")
    assert "$contingency = Get-SimRawCell -Workbook $Workbook" in body
    assert "contingency_ladder" in body
    assert "is not the contingency at the same rung" in body
    assert "[double]$contingency -ne [double]$total" in body, (
        "a run in which the two happened to be equal must not pass silently")
    assert "$total = Get-SimRawCell" not in body, (
        "the total is being read from a raw address rather than the summary block")


def test_34_no_procedure_name_the_contracts_own_is_typed() -> None:
    """The annual endpoint and the four handoff accessors are projected. A typed
    name is a second declaration of a command surface a contract already owns."""
    own = _own_code()
    for projected in ("PCCM_RunAnnualStochastic", "PCCM_AnnualDistributionState",
                      "PCCM_AnnualProfileState", "PCCM_AnnualProfilePx",
                      "PCCM_AnnualYearCount"):
        assert projected not in own, f"{projected} is typed rather than projected"
    assert "$p7.command_surface.annual_endpoint" in own
    assert "$P7.command_surface.handoff_accessors" in own


# ===========================================================================
# E. THE SIX PARTS
# ===========================================================================

def test_40_all_six_parts_are_present_and_in_order() -> None:
    code = _code()
    markers = ["PART 0 - NOTHING HAS RUN", "PART A - A SUCCESSFUL ANNUAL RESULT",
               "PART C - THE ANNUAL ENDPOINT ALONE", "PART D - ITERATIONS CHANGE",
               "PART E - AN INVALID MODEL"]
    positions = [code.index(m) for m in markers]
    assert positions == sorted(positions), markers
    assert "'part B'" in code and "'part 0'" in code


def test_41_the_expected_state_words_are_projected_not_typed() -> None:
    own = _own_code()
    for state in ("'CURRENT'", "'HISTORICAL'", "'OTHER Px'", "'NOT PRODUCED'",
                  '"CURRENT"', '"HISTORICAL"', '"OTHER Px"', '"NOT PRODUCED"'):
        assert state not in own, f"the runner spells the state word {state}"
    for name in ("$notProduced", "$annualCurrent", "$profileCurrent", "$otherPx",
                 "$historical"):
        assert f"{name} = [string]$p7.handoff." in own, name
    # THE MODEL AXIS HAS ITS OWN OWNER. Reading the simulation's CURRENT for a
    # calculation status would merge two vocabularies the contract keeps apart.
    assert "$calcCurrent = [string]$p7.model_states.derived_status[1]" in own


def test_42_part_d_invokes_no_endpoint_and_only_recalculates() -> None:
    """THE VOLATILE ADAPTERS' WHOLE REASON TO EXIST. If an endpoint ran here the
    state would be truthful for a reason that has nothing to do with them."""
    code = _code()
    part = code[code.index("PART D - ITERATIONS CHANGE"):code.index("PART E - AN INVALID MODEL")]
    for endpoint in ("Invoke-Phase6Simulation", "Invoke-P81Endpoint",
                     "Invoke-Phase5ProductionOperation", "PCCM_RunAnnualStochastic"):
        assert endpoint not in part, f"part D invokes {endpoint}"
    assert "Set-NamedValue" in part and "monte_carlo_iterations" in part
    # AND NOTHING OUT-OF-CELL BETWEEN THE INPUT MOVE AND THE FROZEN OBSERVATION.
    # A direct accessor here could persist the derived rows and make the
    # worksheet wrapper look correct on rows this runner had just written.
    window = part[part.index("Set-NamedValue -Workbook $wb -DefinedName $iterationsControl"):
                  part.index("$observationD = Invoke-P81Observation")]
    for out_of_cell in ("$excel.Run(", "Get-P81Handoff", "Invoke-P81Endpoint",
                        "Invoke-P81AccessorParity"):
        assert out_of_cell not in window, (
            f"part D calls {out_of_cell} between the input change and the cell freeze")
    assert "$observationD = Invoke-P81Observation" in part


def test_43_the_derived_rows_are_reported_in_observation_order_and_asserted_nowhere() -> None:
    """THE OBSERVATION THE ROUND EXISTS FOR, AND ITS FOUR OUTCOMES KEPT APART.
    Whether Excel ignored, allowed or refused the write is what is being
    settled; an expectation would decide it in advance, and folding the
    recalculation's effect together with the direct call's would make the
    answer unreadable."""
    body = _function("Invoke-P81Observation")
    assert "THE DERIVED STATUS ROWS, IN OBSERVATION ORDER" in body
    assert "'unchanged'" in body and "'updated'" in body
    assert "recalc: " in body and "direct: " in body, (
        "the two causes of a row moving are not reported apart")
    assert "OBSERVED at " in body
    report = body[body.index("THE DERIVED STATUS ROWS, IN OBSERVATION ORDER"):]
    assert "Add-P81Check" not in report, (
        "the derived rows are asserted rather than reported")


def test_43b_the_two_historical_parts_expect_historical() -> None:
    """PART D IS THE WHOLE REASON THE ADAPTERS ARE VOLATILE. If it expected
    CURRENT it would pass on a sheet that had not noticed the request move at
    all, which is the defect this round corrected."""
    code = _code()
    for marker, ending in (("PART D - ITERATIONS CHANGE", "PART E - AN INVALID MODEL"),
                           ("PART E - AN INVALID MODEL", "[System.GC]::Collect(); ")):
        part = code[code.index(marker):code.index(ending)]
        assert "-Distribution $historical `" in part, (
            f"{marker} does not expect the historical distribution state")
        assert "-Profile $historical" in part, (
            f"{marker} does not expect the historical profile state")
        assert "$annualCurrent" not in part and "$profileCurrent" not in part, (
            f"{marker} expects a current state somewhere")


def test_44_part_e_runs_no_simulation_and_no_annual() -> None:
    code = _code()
    part = code[code.index("PART E - AN INVALID MODEL"):code.index("[System.GC]::Collect(); ")]
    assert "Invoke-Phase6Simulation" not in part
    assert "$p7.command_surface.annual_endpoint" not in part
    assert "Invoke-P81Endpoint -Excel $excel -Endpoint 'PCCM_Calculate'" in part


def test_45_the_selector_parts_prove_the_payload_moved_and_did_not() -> None:
    """Part B must show the sheet unchanged; part C must show it changed. A
    sheet that had cached the old profile would reconcile in both - against the
    wrong total in C - so the payload itself is required to have moved."""
    code = _code()
    part_b = code[code.index("'part B'"):code.index("PART C - THE ANNUAL ENDPOINT ALONE")]
    assert "$null = Add-P81Check 'part B: the persisted run and annual payload are " in part_b
    assert "($movedB.Moved -eq 0)" in part_b
    part_c = code[code.index("PART C - THE ANNUAL ENDPOINT ALONE"):code.index("PART D - ITERATIONS")]
    assert "($movedC.Moved -gt 0)" in part_c, (
        "part C must require the profile to have been republished")
    assert "-ge 0" not in part_c, "a comparison that cannot fail is not a check"


def test_46_the_state_cells_are_cross_checked_against_the_accessors() -> None:
    body = _function("Invoke-P81AccessorParity")
    assert "Get-P81Handoff -Excel $Excel -P7 $P7" in body
    assert "$P7.command_surface.handoff_accessors" in body
    assert "shows what its accessor says" in body
    assert "$Inspection.state.$key.accessor" in body
    # THE COMPARISON IS AGAINST THE FROZEN RECORD, never a fresh read: a cell
    # re-read after the accessor ran would be reporting the probe's own effect.
    assert "$Frozen[$key]" in body
    for reread in ("Get-P81StateCells", "Get-P81Cell", "Invoke-P81Recalculate",
                   "$Excel.Calculate"):
        assert reread not in body, (
            f"the parity comparison {reread}s between the direct call and the check")
    frozen = _function("Invoke-P81FrozenStateChecks")
    assert "$expected = '=' + [string]$Inspection.state.$key.procedure + '()'" in frozen
    assert "-cne $expected" in frozen


def test_46b_every_phase_observes_in_the_one_mandated_order() -> None:
    """ONE FUNCTION OWNS THE ORDER. A phase that assembled the steps for itself
    would be free to assemble them wrongly, which is what happened before this
    function existed."""
    code = _code()
    body = _function("Invoke-P81Observation")
    positions = [body.index(step) for step in (
        "$beforeRecalc = Get-P81DerivedRows",
        "Invoke-P81Recalculate -Excel $Excel -Stage $Stage",
        "$afterRecalc = Get-P81DerivedRows",
        "$frozen = Get-P81StateCells",
        "$evaluated = Add-P81StateEvaluatedCheck",
        "Invoke-P81FrozenStateChecks",
        "Invoke-P81AccessorParity",
        "$afterDirect = Get-P81DerivedRows")]
    assert positions == sorted(positions), (
        "the observation steps are not in the mandated order")
    # THE FROZEN CHECKS MAY NOT REACH A PROCEDURE. A check there that called an
    # accessor would be step 6 wearing step 5's name.
    frozen = _function("Invoke-P81FrozenStateChecks")
    for out_of_cell in ("Get-P81Handoff", "$Excel", "Application.Run", ".Run("):
        assert out_of_cell not in frozen, f"the frozen checks reach {out_of_cell}"
    # AND EVERY PHASE GOES THROUGH IT. No phase may freeze cells for itself.
    for stage in ("part 0", "part A", "part B", "part C", "part D", "part E"):
        assert f"-Stage '{stage}'" in code, stage
    assert code.count("Invoke-P81Observation -Excel $excel") == 6
    assert code.count("Get-P81StateCells") == 2, (
        "the state cells are read outside the observation orchestrator")
    # Definition, the parity step, and the failed-gate diagnostic. Nowhere else.
    assert code.count("Get-P81Handoff") == 3, (
        "the accessors are called outside the parity step and the failed-gate diagnostic")


def test_47_the_empty_part_refuses_a_fabricated_zero_and_a_blank_pass() -> None:
    code = _code()
    part = code[code.index("PART 0 - NOTHING HAS RUN"):code.index("PART A - A SUCCESSFUL")]
    assert "the annual table shows nothing, not zeros" in part
    assert "the reconciliation reports nothing at all, and passes nothing" in part
    assert "$p8.annual.row_window" in part, "the far end of the window is inspected too"


# ===========================================================================
# F. LIFECYCLE
# ===========================================================================

def test_50_the_com_lifecycle_is_the_accepted_one() -> None:
    code = _code()
    assert code.count("New-ReleaseLedger") == 1
    assert code.index("New-ReleaseLedger") < code.index("New-Object -ComObject Excel.Application")
    assert code.count("$comAcquired = $comAcquired + 1") == 3
    assert "$wb.Close($false)" in code and "$excel.Quit()" in code
    assert "$Error.Clear()" in code
    assert "Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90" in code
    assert ".Save()" not in code and "SaveAs" not in code
    for claim in ("the owned Excel process exited naturally",
                  "no emergency cleanup was required",
                  "every COM object this runner acquired was released",
                  "every COM release succeeded",
                  "every COM release left 0 outstanding references",
                  "every transient release inside the accepted readers succeeded"):
        assert f"'{claim}'" in code, claim


def test_51_the_tree_is_proved_clean_before_excel_is_started() -> None:
    code = _code()
    assert code.index("Get-P81SourceRevision -RepoRoot $repoRoot") < code.index(
        "New-Object -ComObject Excel.Application")


def test_52_the_report_is_written_through_and_separates_prerequisites() -> None:
    code = _code()
    assert "Set-Content -LiteralPath $script:P81Path" in _function("Write-P81Line")
    assert "$Kind = 'RESULT'" in _function("Add-P81Check")
    assert "($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0)" in code
    assert "($results.Count -gt 0)" in code


def test_53_the_recalculation_records_the_calculation_mode() -> None:
    """A state that failed to update because the workbook was left in manual is
    a different finding from one that failed because the adapter did not fire."""
    body = _function("Invoke-P81Recalculate")
    assert "$Excel.Calculation" in body and "$Excel.Calculate()" in body
    assert "Application.Calculation = " in body


# ===========================================================================
# G. THE SCALAR-VERSUS-ARRAY DEFECT THE FIRST WINDOWS RUN FOUND
# ===========================================================================

def test_60_no_comma_expression_is_an_operand_of_arithmetic() -> None:
    """THE EXACT FAILURE, AS A CLASS. Not the one line: any place where a comma
    list and arithmetic meet at the same level."""
    found = _comma_bound_arithmetic(_text())
    assert not found, (
        "a comma expression is an operand of arithmetic; PowerShell's comma binds "
        "tighter, so this is an Object[] in an arithmetic position:\n  " +
        "\n  ".join(f"line {line}: {text}" for line, text in found))


def test_61_the_detector_detects(caplog=None) -> None:
    """A CONTROL THAT CANNOT FAIL IS NOT A CONTROL. The detector is shown to
    fire on the expression that actually broke the run, and on the one other
    instance of the class in the tree - which is in the FROZEN harness, is
    genuinely broken, and is left exactly where history put it."""
    broken = "$fabricated = @(0, 1, [int]$p8.annual.row_window - 1)\n"
    assert _comma_bound_arithmetic(broken), "the detector misses the defect it was written for"
    safe = "$last = [int]$p8.annual.row_window - 1\n$fabricated = @(0, 1, $last)\n"
    assert not _comma_bound_arithmetic(safe), "the detector fires on the correction"
    call = "$x = [Math]::Min($YearCount + 2, [int]$window)\n"
    assert not _comma_bound_arithmetic(call), (
        "the detector fires on a method call, whose arguments are not a comma expression")
    parenthesised = "$x = @(0, 1, ([int]$w - 1))\n"
    assert not _comma_bound_arithmetic(parenthesised)


def test_62_no_live_windows_runner_carries_the_defect() -> None:
    """AND IT IS SWEPT, not checked one runner at a time. The frozen acceptance
    harness carries the only other instance in the tree - `$Block[$Year, $Offset
    + 1]`, which really does fail with "cannot index into a 2 dimensional array
    with index [1,0,1]" - and it is history: never patched, never run again, and
    named here so its exclusion is a statement rather than a gap."""
    frozen = WINDOWS / "phase7_acceptance_scenarios.ps1"
    for path in sorted(WINDOWS.glob("*.ps1")):
        found = _comma_bound_arithmetic(path.read_text(encoding="utf-8"))
        if path == frozen:
            assert found, "the frozen harness no longer carries its known instance"
            assert len(found) == 1, found
            continue
        assert not found, f"{path.name} carries comma-bound arithmetic: {found}"


def test_63_the_failing_part_0_path_is_covered_and_corrected() -> None:
    """THE EXACT LINE THAT FAILED, and the shape that replaced it: the far end
    of the window is a named integer computed before the list, not arithmetic
    inside it."""
    code = _code()
    part = code[code.index("$observation0 = Invoke-P81Observation"):
                code.index("PART A - A SUCCESSFUL ANNUAL RESULT")]
    assert "$lastAnnualOffset = [int]$p8.annual.row_window - 1" in part, (
        "the far end of the annual window is not computed into a named integer")
    assert "foreach ($offset in @(0, 1, $lastAnnualOffset)) {" in part
    assert "@(0, 1, [int]$p8.annual.row_window - 1)" not in code, (
        "the expression that broke the first Windows run is back")
    # AND THE SAMPLE IS STILL THE ONE PART 0 NEEDS: the first row, its
    # neighbour, and the far end of the structural window.
    assert "$p8.annual.row_window" in part
    assert _projection()["annual"]["row_window"] >= 200


def test_64_projected_row_arithmetic_is_always_cast_to_an_integer() -> None:
    """A ROW NUMBER FROM JSON IS AN Int64, AND A COLUMN IS A STRING. Every place
    the runner does arithmetic on a projected coordinate casts it first, so an
    operand can never arrive as whatever ConvertFrom-Json felt like returning."""
    own = _own_code()
    for projected in re.finditer(
            r"\$(?:p8|Inspection|P7|simInspection|SimInspection)\."
            r"[\w.$]*(?:row|rows|first_row|header_row|row_window|Count)\b\s*[-+]\s",
            own):
        text = own[max(0, projected.start() - 12):projected.end()]
        # A CAST OR A COUNT. `[int]` for arithmetic and `[string]` for the
        # concatenations that build an address are both scalar conversions; what
        # must never happen is a raw ConvertFrom-Json member in an operand.
        assert re.search(r"\[(?:int|string|double|long)\]", text) or ".Count" in text, (
            f"projected row arithmetic on an uncast operand: {text!r}")


def test_65_helpers_that_return_a_collection_are_consumed_as_one() -> None:
    """MADE EXPLICIT, so a collection is never read as a scalar. These three
    return sets, not values, and every call site takes them as sets."""
    text = _text()
    for name, returns in (("Get-P81RowsInGroup", "@($out)"),
                          ("Get-P81DerivedRows", "an ordered dictionary"),
                          ("Get-P81StateCells", "an ordered dictionary")):
        assert name in text, name
    assert "return @($out)" in _function("Get-P81RowsInGroup"), (
        "Get-P81RowsInGroup must return a collection explicitly")
    # AND NOBODY SUBTRACTS FROM ONE.
    for call in re.finditer(r"(Get-P81RowsInGroup|Get-P81DerivedRows|Get-P81StateCells)"
                            r"[^\n]*", text):
        line = call.group(0)
        assert not re.search(r"\)\s*[-+*/]\s", line), (
            f"a collection-returning helper is used in arithmetic: {line.strip()!r}")


# ===========================================================================
# THE LOCKED FX SEED - THE PREREQUISITE THE FIRST RUN NEVER PRIMED
# ===========================================================================
# WHAT HAPPENED. The second Windows run passed the whole of Part 0 and then died
# on the first statement of Part A with the accepted helper's own refusal:
#
#     the locked FX seed was never captured. Save-Phase5LockedFxSeed must run on
#     the untouched Stage-B workbook, before any Phase-5 mutation.
#
# Step C of Invoke-Phase5FixtureSteps restores the FX table from a seed captured
# before step A, Clear-Phase5Registers, wiped the registers. P8-1 was composed
# from W8's parts and that one line was dropped, so there was nothing to restore
# from. The controls below pin the line, its position, and the two properties
# that make the position legal: the capture is a read, and Part 0 is a read.
#
# THE ORDERING RULES ARE FUNCTIONS, NOT ASSERTIONS BURIED IN A TEST, so the
# mutation control at the end can run the very same rules against a deliberately
# broken copy of the source and require them to refuse it.

# EVERY WRITE THE RUNNER COULD MAKE. A Phase-5 mutation is any of these; the
# fixture's own first mutation, Clear-Phase5Registers, is reached through
# Set-Phase5Fixture.
MUTATORS = (
    "Set-Phase5Fixture", "Clear-Phase5Registers", "Set-NamedValue", "Set-TableCell",
    "Set-P81NamedText", "Invoke-P81Endpoint", "Add-BlankTableRow", "Remove-TableRow",
    "Set-Phase5TypedCell", "Reset-Phase5FxTable", "ClearContents", ".Value2 =",
)
CAPTURE = "$null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection"
SESSION = "$rel = New-ReleaseLedger"
COMPILED = "the current VBAProject compiles in real Excel"
PART0 = "PART 0 - NOTHING HAS RUN"
PARTA = "PART A - A SUCCESSFUL ANNUAL RESULT"


def _capture_is_singular(code: str) -> None:
    """EXACTLY ONE CAPTURE, in the accepted call shape."""
    count = code.count("Save-Phase5LockedFxSeed")
    assert count == 1, f"the locked FX seed must be captured exactly once; found {count}"
    assert CAPTURE in code, "the capture is not the call shape W2 through W8 use"


def _capture_is_at_the_accepted_point(code: str) -> None:
    """AFTER THE COMPILE PREREQUISITE, BEFORE PART 0, BEFORE THE FIXTURE."""
    assert code.index(COMPILED) < code.index("Save-Phase5LockedFxSeed"), (
        "the seed is captured before the compile prerequisite")
    assert code.index("Save-Phase5LockedFxSeed") < code.index(PART0), (
        "the seed is captured after Part 0 has begun")
    assert code.index("Save-Phase5LockedFxSeed") < code.index("Set-Phase5Fixture"), (
        "the seed is captured after the fixture has already mutated the workbook")


def _nothing_mutates_before_the_capture(code: str) -> None:
    """THE HELPER'S OWN PRECONDITION: an untouched workbook."""
    before = code[code.index(SESSION):code.index("Save-Phase5LockedFxSeed")]
    for mutator in MUTATORS:
        assert mutator not in before, (
            f"{mutator} runs before the locked FX seed is captured")


def _part_0_completes_before_any_mutation(code: str) -> None:
    """PART 0 RUNS TO ITS END, AND WRITES NOTHING, before the fixture exists."""
    part0, fixture = code.index(PART0), code.index("Set-Phase5Fixture")
    assert part0 < code.index(PARTA) < fixture, (
        "Part A does not begin between the end of Part 0 and the fixture")
    body = code[part0:fixture]
    for mutator in MUTATORS:
        assert mutator not in body, f"Part 0 mutates the workbook through {mutator}"


def test_70_the_locked_fx_seed_is_captured_exactly_once_by_the_accepted_helper() -> None:
    """CAPTURED, ONCE, AND NOT REINVENTED. The seed comes from the accepted
    Phase-5 helper the fixture itself reads back from - not from a second
    capture, and not from a local copy of the logic."""
    code = _code()
    _capture_is_singular(code)
    assert "Save-Phase5LockedFxSeed" in PHASE5.read_text(encoding="utf-8"), (
        "the accepted helper is not where the runner dot-sources it from")
    for name in ("Save-Phase5LockedFxSeed", "Get-Phase5LockedFxSeed", "Reset-Phase5FxTable"):
        assert not re.search(rf"^function\s+{name}\b", code, re.M), (
            f"{name} is redefined in the runner instead of reused")
    assert "Phase5LockedFxSeed = " not in code, (
        "the runner assigns the accepted helper's seed variable itself")


def test_71_the_capture_is_at_the_lifecycle_point_the_accepted_runners_use() -> None:
    """THE ORDER W2 THROUGH W8 WERE ACCEPTED WITH, proved against W8 rather than
    described: compile prerequisite, then the seed, then the fixture."""
    _capture_is_at_the_accepted_point(_code())
    w8 = accepted._ps_code(W8)
    assert (w8.index(COMPILED) < w8.index("Save-Phase5LockedFxSeed")
            < w8.index("Set-Phase5Fixture")), "W8 is not the order this control claims"


def test_72_nothing_mutates_the_phase_5_fixture_before_the_seed_is_captured() -> None:
    """A SEED CAPTURED AFTER A WRITE IS THE WRONG SEED, and the helper cannot
    tell. Nothing in the session may write before this line."""
    _nothing_mutates_before_the_capture(_code())


def test_73_part_0_completes_before_the_first_fixture_mutation() -> None:
    """PART 0 IS AN OBSERVATION OF A WORKBOOK NOTHING HAS TOUCHED. It runs to its
    end before the fixture is applied, and it writes nothing itself."""
    code = _code()
    _part_0_completes_before_any_mutation(code)
    # AND THE GATE IS STILL INSIDE PART 0, ahead of everything the fixture makes
    # possible: a run that cannot evaluate the four cells stops before it writes.
    body = code[code.index(PART0):code.index("Set-Phase5Fixture")]
    assert body.index("$observation0 = Invoke-P81Observation") < body.index(
        "$stoppedOnStateEvaluation = $true")


def test_74_the_capture_is_read_only_so_part_0_still_sees_an_untouched_workbook() -> None:
    """THE PROPERTY THAT MAKES THE POSITION LEGAL. Save-Phase5LockedFxSeed sits
    ahead of Part 0, which would be indefensible if it wrote anything. Read the
    accepted helper and prove it does not: it is handed no Excel application, and
    everything it touches is a reader."""
    phase5 = accepted._ps_code(PHASE5)
    match = re.search(r"^function\s+Save-Phase5LockedFxSeed\s*\{", phase5, re.M)
    assert match, "Save-Phase5LockedFxSeed is not defined in the accepted Phase-5 source"
    start, depth, body = match.end() - 1, 0, None
    for index in range(start, len(phase5)):
        if phase5[index] == "{":
            depth += 1
        elif phase5[index] == "}":
            depth -= 1
            if depth == 0:
                body = phase5[start:index]
                break
    assert body, "the helper body could not be read"
    for writer in ("Set-TableCell", "Set-Phase5TypedCell", "Set-NamedValue",
                   "Add-BlankTableRow", "Remove-TableRow", "ClearContents",
                   ".Value2 =", ".Formula", "$Excel", "Excel.Run"):
        assert writer not in body, f"the capture is not read-only: it uses {writer}"
    assert "Get-Phase5TypedTableBody" in body, "the capture no longer reads the FX table"


def test_75_the_fixture_is_applied_through_the_proven_helper_in_its_own_order() -> None:
    """NO STEP OF THE ACCEPTED FIXTURE IS RESEQUENCED HERE. The runner asks for
    the whole fixture, once, with W8's argument shape; the A-to-H order stays
    where it was accepted, inside Invoke-Phase5FixtureSteps, which is also where
    the seed is read back."""
    code = _code()
    assert code.count("Set-Phase5Fixture") == 1, "the fixture is applied more than once"
    call = re.search(r"Set-Phase5Fixture(?:.|\n)*?-Model \$model\)", code)
    assert call, "the fixture call does not have the accepted argument shape"
    for argument in ("-Excel $excel", "-Workbook $wb", "-Manifest $manifest",
                     "-Inspection $inspection", "-Model $model"):
        assert argument in call.group(0), f"the fixture call is missing {argument}"
    phase5 = accepted._ps_code(PHASE5)
    assert re.search(r"^function\s+Invoke-Phase5FixtureSteps\b", phase5, re.M), (
        "the accepted fixture step function is gone")
    assert "Get-Phase5LockedFxSeed" in phase5, "the fixture no longer reads the seed back"
    for step in ("Clear-Phase5Registers", "Reset-Phase5FxTable", "Invoke-Phase5FixtureSteps"):
        assert step not in code, f"the runner reaches around the fixture and calls {step}"


@pytest.mark.parametrize("name,mutate", [
    # THE CAPTURE MOVED PAST THE FIXTURE - the defect the second run died of,
    # reintroduced one statement later instead of omitted.
    ("the capture moved after the fixture",
     lambda code: code.replace(CAPTURE + "\n", "", 1).replace(
         "$applied = [string](Set-Phase5Fixture",
         "$applied = [string](Set-Phase5Fixture", 1).replace(
         "-Model $model)", "-Model $model)\n        " + CAPTURE, 1)),
    # THE CAPTURE DELETED - the state the runner was actually in.
    ("the capture deleted", lambda code: code.replace(CAPTURE + "\n", "", 1)),
    # THE CAPTURE DUPLICATED - a second read after the fixture would overwrite
    # the untouched seed with a mutated one, silently.
    ("the capture duplicated",
     lambda code: code.replace("-Model $model)", "-Model $model)\n        " + CAPTURE, 1)),
    # THE FIXTURE MOVED AHEAD OF PART 0 - Part 0 stops being an observation of a
    # workbook nothing has touched.
    ("the fixture moved ahead of part 0",
     lambda code: code.replace(
         CAPTURE, CAPTURE + "\n    $null = Set-Phase5Fixture -Excel $excel", 1)),
])
def test_76_the_four_orderings_that_would_break_the_prerequisite_are_refused(
        name: str, mutate) -> None:
    """THE MUTATIONS, RUN AGAINST THE RULES ABOVE. Each one is applied to a copy
    of the source in memory - nothing on disk changes - and at least one of the
    four ordering rules must refuse it. A mutation that survives every rule means
    the rules are decoration."""
    mutated = mutate(_code())
    assert mutated != _code(), f"the mutation '{name}' changed nothing"
    refused = []
    for rule in (_capture_is_singular, _capture_is_at_the_accepted_point,
                 _nothing_mutates_before_the_capture, _part_0_completes_before_any_mutation):
        try:
            rule(mutated)
        except (AssertionError, ValueError) as failure:
            refused.append(f"{rule.__name__}: {failure}")
    assert refused, f"'{name}' survived every ordering rule"
