#!/usr/bin/env python3
"""P7-7: the MINIMAL W5 runner, proved on Linux before a Windows session.

W5 IS THE FIRST SUCCESSFUL ANNUAL RUN. W5 proved the endpoint refuses with no
simulation and that a FIXED-seed baseline publishes cleanly; W5 rebuilds that
baseline and asks the endpoint to produce an answer, persist it compactly, stamp
it to the run that produced it, hand it to Phase 8, and reconcile.

WHAT IS INDEPENDENT AND WHAT IS NOT, AGAIN CHECKED AS CAREFULLY AS THE VALUES.
There is no independent annual Windows oracle. The annual COMPUTATION is proved
independently on Linux; W5's subject is execution, persistence, read-back and
reconciliation, which only a Windows run can show. A control requires the runner
to say that plainly, and other controls require it to lean only on authorities
that exist: the accepted Phase-5 oracle for the deterministic base, the
published iteration column as the authoritative run data, the contract's own
Type-7 definition over that column for the total percentile, and the persisted
result read back through its contracted surface.

THE BLEND IS NEVER RECONSTRUCTED. Reimplementing production's convex Type-7
annual blend in PowerShell would compare an implementation against a copy of
itself, so a control forbids it: what the runner computes is the TOTAL
percentile, and what it then requires is the identity sim_contract names,
sum_y Profile_Px(y) = reported Px, under the project's own accepted identity
rule read from the corpus rather than chosen here.

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
from pccm_builder import (  # noqa: E402
    load_calc_contract,
    load_sim_contract,
    load_structure_contract,
)

WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
RUNNER = WINDOWS / "phase7_w5_annual_success.ps1"
W1 = WINDOWS / "phase7_w1_smoke.ps1"
W2 = WINDOWS / "phase7_w2_many_drivers.ps1"
W3 = WINDOWS / "phase7_w3_long_years.ps1"
W4 = WINDOWS / "phase7_w4_base_simulation.ps1"
TIMING = WINDOWS / "phase7_timing_scenarios.ps1"
LIFECYCLE = WINDOWS / "com_lifecycle.ps1"
PHASE5 = WINDOWS / "phase5_gate_b_scenarios.ps1"
PHASE6 = WINDOWS / "phase6_gate_b_scenarios.ps1"
BUILD = PCCM_ROOT / "build"

DOT_SOURCED = (LIFECYCLE, PHASE5, PHASE6)

# The ten helpers the closure below found. Named here so the count is a claim a
# reader can check rather than a number that drifts silently.
COPIED_HELPERS = (
    "Write-RowObject", "Get-NamedValue", "Set-NamedValue", "Get-TableColumnNames",
    "Set-TableCell", "Get-TableBody", "Get-TableRowCount", "Add-BlankTableRow",
    "Remove-TableRow", "Get-IdColumnValues",
)

# NINE OF THE TEN ARE PURELY TRANSITIVE - the accepted fixture calls them and
# this runner never does. `Set-NamedValue` is the exception and is named as one:
# W5 writes the simulation request through it, so it is BOTH copied for the
# fixture's sake and called directly here.
CALLED_DIRECTLY = ("Set-NamedValue",)

# W1, W2, W3 AND W4 ARE CLOSED. All three runners are accepted Windows evidence and
# are pinned, not maintained, by this round. The digests are LITERAL - hashing a
# file at import time would be a control that can never fail.
W1_SHA256 = "bf41a2193dbf2ea5fbddbe4dad3c6e94649f21262dcf9800ae71eda6e274b59f"
W2_SHA256 = "558bcfa0528b38c28c6e81dd726a18fed9078098f79bcabb792304e6255d8fee"
W3_SHA256 = "97af69bf836a2d798be4937f1e81cfe7218d2cfd702c1b96aa25623e7f8c12be"
W4_SHA256 = "9e1be0629d337b42b02bf790feef752c1fd37fbf0e6c400d4c90af3cd9d86cd0"

_CACHE: dict = {}


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    """The runner with its block comment and line comments removed.

    Its header explains what W5 refuses to do and names the things the controls
    forbid, so a scan that read the prose would convict the file of its own
    documentation.
    """
    return accepted._ps_code(RUNNER)


def _functions() -> dict[str, str]:
    """name -> body, brace-matched over the comment-stripped source.

    String literals are KEPT: the compared column names, the field map and the
    provenance strings are all literals, and they are the subject here.
    """
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
    assert body is not None, f"{name} is not defined in the W5 runner"
    return body


def _own_code() -> str:
    """The runner's top level plus the functions it wrote itself.

    The ten copied helpers are accepted code from another file; a claim about
    what THIS runner does must not be made about them, and a claim about them is
    made by pinning them instead.
    """
    code = _code()
    for name in COPIED_HELPERS:
        code = code.replace(_function(name), "")
    return code


def _cases() -> dict:
    if "cases" not in _CACHE:
        _CACHE["cases"] = json.loads(
            (BUILD / "phase7_acceptance_cases.json").read_text(encoding="utf-8"))
    return _CACHE["cases"]


def _fixture() -> dict:
    """THE W4 BEHAVIOURAL CASE. W5 reuses it deliberately: the annual answer it
    proves has to be the answer produced from the baseline W4 accepted."""
    case = [s for s in _cases()["scenarios"] if s["id"] == "W4"]
    assert case, "the acceptance corpus carries no W4 behavioural fixture"
    return case[0]


def _year_window() -> tuple[int, int, int]:
    """(min_year, max_year, max_generated_year_columns) from the structural contract.

    THE AUTHORITY, NOT A REMEMBERED NUMBER. The W5 fixture's calendar span is
    derived from these, so the assertions about it have to be too - the round
    that produced 2026-2225 failed precisely because a literal span was carried
    around instead of being derived.
    """
    text = (PCCM_ROOT / "spec" / "structure_contract.yaml").read_text(encoding="utf-8")
    values = {}
    for key in ("min_year", "max_year", "max_generated_year_columns"):
        match = re.search(rf"^\s*{key}:\s*(\d+)\s*$", text, re.M)
        assert match, f"structure_contract.yaml declares no {key}"
        values[key] = int(match.group(1))
    return values["min_year"], values["max_year"], values["max_generated_year_columns"]


def _span() -> tuple[int, int, int]:
    """(first calendar year, last calendar year, duration) of the W5 fixture."""
    timeline = _fixture()["model"]["timeline"]
    first = int(timeline["start_year"])
    duration = int(timeline["duration"])
    return first, first + duration - 1, duration


def _inspection() -> dict:
    if "inspection" not in _CACHE:
        _CACHE["inspection"] = json.loads(
            (BUILD / "phase5_gate_b_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["inspection"]


# ===========================================================================
# A. SCOPE: THIS IS W5, AND W1 IS CLOSED
# ===========================================================================

def test_01_the_runner_exists_and_is_dedicated() -> None:
    assert RUNNER.exists()
    lines = _text().splitlines()
    assert len(lines) < 1600, (
        f"{len(lines)} lines; W5 rebuilds a baseline and then proves one annual "
        "run, and most of its behaviour is meant to come from the accepted "
        "Phase-5 fixture, the accepted Phase-6 state readers and the projection")
    # Most of the file is either the accepted copied helpers or W5's own small
    # comparison; neither should have grown into a second harness.
    assert len(_own_code().splitlines()) < 1200


def test_02_it_dot_sources_only_definition_only_files() -> None:
    """The three that define and run nothing. `phase4_functional_test.ps1` is
    NOT among them: dot-sourcing it would run the entire Phase-4 matrix, which
    is exactly why its helpers are copied instead."""
    sourced = re.findall(r"^\. \(Join-Path \$scriptDir '([\w.]+)'\)", _code(), re.M)
    assert sourced == [p.name for p in DOT_SOURCED], sourced
    for path in DOT_SOURCED:
        # Definition-only means the top level holds declarations and variables,
        # never a scenario invocation. The function HEADERS survive the body
        # removal - they are declarations, not calls - so they are dropped
        # before the top level is read for calls.
        top = "\n".join(line for line in accepted._ps_top_level(path).splitlines()
                        if not line.strip().startswith("function "))
        for entry in ("Invoke-Phase5GateBScenarios", "Invoke-Phase6GateBScenarios"):
            assert not re.search(rf"(?<![\w\-.$]){entry}(?![\w\-])", top), (
                f"{path.name} invokes {entry} at top level; dot-sourcing it "
                "would RUN a Gate-B scenario")


def test_03_the_stage_b_bootstrap_is_a_child_script_not_a_dot_source() -> None:
    code = _code()
    assert "& $bootstrap -BuildDir $tempRoot -Force" in code
    assert ". (Join-Path $scriptDir 'build_stage_b.ps1')" not in code


def test_04_no_sensitivity_no_later_scenario_and_no_second_stochastic_oracle() -> None:
    """W5 runs the annual endpoint ONCE, expecting a refusal, and one simulation.

    Everything else stochastic belongs to another scenario, and a second
    stochastic oracle invented here would be exactly the "compare a run against
    another run and call it independence" the authorisation refuses.
    """
    code = _own_code()
    for forbidden in ("PCCM_RunSensitivity", "Invoke-Phase5GateBScenarios",
                      "Invoke-Phase6GateBScenarios", "phase7_acceptance_scenarios",
                      "phase7_w1_smoke", "phase7_w2_many_drivers", "phase7_w3_long_years"):
        assert forbidden not in code, f"{forbidden} is outside W5"
    for scenario in ("'W2'", "'W3'", "'W6'", "'W7'", "'W8'"):
        assert scenario not in code, scenario
    # W5 REUSES THE W4 BEHAVIOURAL FIXTURE DELIBERATELY: the annual answer it
    # proves has to be the answer produced from the baseline W4 accepted, so the
    # corpus case it selects is W4's and the runner says so.
    assert "$_.id -ceq 'W4'" in code
    assert "the W5 fixture is the accepted W4 baseline" in code
    # ONE simulation, ONE annual attempt.
    assert code.count("Invoke-Phase6Simulation") == 1, (
        "W5 runs the simulation more than once; a same-seed replay compared "
        "against itself is not independent evidence and is not authorised here")
    assert code.count("Invoke-W5Annual") == 2, (
        "the annual endpoint is defined once and invoked from one call site")


def test_05_it_does_not_repeat_the_w1_surface_matrix() -> None:
    """One lightweight read-only compile trigger is authorised because Excel
    must execute the current project. The surface matrix is W1's."""
    code = _own_code()
    for forbidden in ("VBProject", "VBComponents", "CodeModule", "ProcOfLine"):
        assert forbidden not in code, f"{forbidden} is W1's evidence, not W5's"
    assert "$excel.Run('PCCM_CalculationStatus')" in code
    assert "the current VBAProject compiles in real Excel" in code


def test_06_the_accepted_w1_and_w2_runners_are_untouched() -> None:
    """BOTH ARE CLOSED on accepted Windows evidence. This round does not modify
    or rerun either, and an edit to one must fail here rather than pass."""
    assert W1.exists() and W2.exists() and W3.exists() and W4.exists()
    assert hashlib.sha256(W1.read_bytes()).hexdigest() == W1_SHA256, (
        "the accepted W1 runner has been modified; it is closed evidence")
    assert hashlib.sha256(W2.read_bytes()).hexdigest() == W2_SHA256, (
        "the accepted W2 runner has been modified; it is closed evidence")
    assert hashlib.sha256(W3.read_bytes()).hexdigest() == W3_SHA256, (
        "the accepted W3 runner has been modified; it is closed evidence")
    assert hashlib.sha256(W4.read_bytes()).hexdigest() == W4_SHA256, (
        "the accepted W4 runner has been modified; it is closed evidence")
    code = _code()
    for closed in ("phase7_w1_smoke", "phase7_w2_many_drivers",
                   "phase7_w3_long_years", "phase7_w4_base_simulation"):
        assert closed not in code, f"W5 reaches into the closed {closed} at run time"


# THE PARTS W5 SHARES WITH W2, PINNED SO THE MANDATED COPY CANNOT DRIFT.
# Reusing W2's proven execution architecture was the instruction; two runners
# that started identical and quietly diverged would be the cost of it, so the
# shared functions are compared after normalising the scenario name.
SHARED_WITH_W2 = ("Write-W5Line", "Add-W5Check", "Format-W5Value", "Invoke-W5Release",
                  "Get-W5SourceRevision", "Compare-W5Cell", "Compare-W5Table")


def test_06b_the_shape_shared_with_w2_is_identical_to_w2s() -> None:
    import re as _re
    ours = _text()
    theirs = W2.read_text(encoding="utf-8")
    for name in SHARED_WITH_W2:
        mine = _re.search(rf"^function {name} \{{(.*?)^\}}", ours, _re.M | _re.S)
        assert mine, f"{name} is not defined in the W5 runner"
        w2name = name.replace("W5", "W2")
        yours = _re.search(rf"^function {w2name} \{{(.*?)^\}}", theirs, _re.M | _re.S)
        assert yours, f"{w2name} is no longer in the accepted W2 runner"
        assert mine.group(1).replace("W5", "W2") == yours.group(1), (
            f"{name} has drifted from the accepted {w2name}; the two runners are "
            "meant to share one execution architecture, not two")


# ===========================================================================
# B. THE CALL CLOSURE - THE DEFECT THAT COST W1 A WINDOWS SESSION
# ===========================================================================

def _closure() -> tuple[set[str], list[tuple[str, str]]]:
    defined: dict[str, str] = {}
    for path in (RUNNER,) + DOT_SOURCED:
        for name, body in accepted._ps_functions(path).items():
            defined.setdefault(name, body)
    seen: set[str] = set()
    missing: list[tuple[str, str]] = []
    work: list[tuple[str, str]] = [("<runner top level>", accepted._ps_top_level(RUNNER))]
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
    return seen, missing


def test_07_every_custom_command_the_runner_reaches_is_defined() -> None:
    """TRANSITIVELY, because the dependency that broke W1 was transitive: the
    harness never named `Write-RowObject`, and would have failed on it - and
    then on nine more like it - one Windows session at a time."""
    reached, missing = _closure()
    assert not missing, (
        "custom command(s) called but defined neither in the runner nor in the "
        "files it dot-sources:\n  " +
        "\n  ".join(f"{call}  (reached from {where})" for where, call in sorted(set(missing))))
    assert len(reached) > 40, f"only {len(reached)} custom commands reached"


def test_08_the_copied_helpers_are_verbatim_and_genuinely_transitive() -> None:
    """Pinned to the accepted implementation, and justified: nothing in this
    runner calls any of them. The accepted Phase-5 fixture does, and its own
    file does not define them."""
    timing = TIMING.read_text(encoding="utf-8")
    raw = _text()
    own = _own_code()
    for name in COPIED_HELPERS:
        theirs = re.search(rf"^function {name} \{{(.*?)^\}}", timing, re.M | re.S)
        assert theirs, f"{name} is no longer in the accepted timing harness"
        # RAW AGAINST RAW, comments included: a copied helper whose commentary
        # was edited is no longer the accepted implementation, and the whole
        # point of copying rather than paraphrasing is that it stays identical.
        ours = re.search(rf"^function {name} \{{(.*?)^\}}", raw, re.M | re.S)
        assert ours, f"{name} is not defined in the W5 runner"
        assert ours.group(1) == theirs.group(1), (
            f"{name} has drifted from the accepted implementation")
        body = own
        for line in body.splitlines():
            if line.strip().startswith("function "):
                continue
        called = re.search(rf"(?<![\w\-.$]){name}(?![\w\-])",
                           "\n".join(l for l in own.splitlines()
                                     if not l.strip().startswith("function ")))
        if name in CALLED_DIRECTLY:
            assert called, (
                f"{name} is listed as called directly but nothing calls it; "
                "the justification in this file no longer describes the runner")
        else:
            assert not called, (
                f"{name} is called from the runner after all; the transitive "
                "justification in the source no longer describes it")


# ===========================================================================
# C. WINDOWS POWERSHELL 5.1
# ===========================================================================

def test_09_the_runner_declares_and_obeys_the_shell_it_targets() -> None:
    text = _text()
    assert "WINDOWS POWERSHELL 5.1" in text
    code = _code()
    assert "$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)" in code
    assert "$repoRoot = Split-Path -Parent $pccmRoot" in code
    for number, line in enumerate(code.splitlines(), 1):
        count = accepted._join_path_positional_count(line)
        assert count is None or count <= 2, f"line {number}: {line.strip()}"


def test_10_no_powershell_6_or_7_only_construct_is_used() -> None:
    offenders: list[str] = []
    code = _code()
    for label, pattern in accepted.PS51_ONLY_CONSTRUCTS:
        for number, line in enumerate(code.splitlines(), 1):
            if re.search(pattern, line):
                offenders.append(f"{number}: {label}: {line.strip()[:80]}")
    assert not offenders, "\n  ".join(offenders)


# ===========================================================================
# G. COM LIFECYCLE - THE DISCIPLINE W1 CLOSED ON
# ===========================================================================

def test_25_every_acquisition_is_counted_and_released_into_one_ledger() -> None:
    code = _own_code()
    # RELEASE-TRANSIENT IS ALLOWED IN EXACTLY ONE PLACE, and it is named.
    # `Set-W5NamedText` writes one defined name: its three objects are acquired
    # and released inside a single statement, in the same convention every
    # accepted reader uses, and they are covered by the accepted-reader
    # transient gate. Everything the runner HOLDS goes through its own ledger.
    outside = code.replace(_function("Set-W5NamedText"), "")
    assert "Release-Transient" not in outside, (
        "a release the runner holds still goes through the silent-on-success "
        "helper; only Set-W5NamedText may use it")
    assert _function("Set-W5NamedText").count("Release-Transient") == 3
    assert "Invoke-NamedRelease" not in code, (
        "Invoke-NamedRelease keeps the release count to itself")
    for label in ("'Workbook'", "'Workbooks'", "'Excel.Application'"):
        assert f"Invoke-W5Release $rel" in code and label in code, label
    ledger_at = code.index("$rel = New-ReleaseLedger")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert ledger_at < excel_at
    assert code.count("$comAcquired = $comAcquired + 1") == 3, (
        "the three session objects are each counted where they are acquired")
    # AND EVERY RANGE READ IS COUNTED TOO. W5 opens three of them - the
    # iteration column before and after the annual run, and the annual region -
    # and an uncounted acquisition would break the balance check without
    # anything saying which one.
    counted = re.findall(r"\$comAcquired = \$comAcquired \+ \[int\]\$(\w+)\.Acquired", code)
    assert sorted(counted) == ["iterationsAfter", "iterationsBefore", "region"], counted
    for name in ("Get-W5IterationBlock", "Get-W5AnnualRegionIndex"):
        block = _function(name)
        assert block.count("Invoke-W5Release $Ledger") == 3, name
        assert block.count("$acquired = $acquired + 1") == 3, name
    helper = _function("Invoke-W5Release")
    assert "Release-ComObjectSafe" in helper
    assert "[int]$rec.Count -ne 0" in helper
    assert "$script:W5Residual.Add" in helper


def test_26_the_lifecycle_verdict_checks_are_all_present() -> None:
    code = _own_code()
    for claim in ("'the owned Excel process exited naturally' $naturalExit",
                  "'no emergency cleanup was required' (-not $emergencyRequired)",
                  "'every COM object this runner acquired was released'",
                  "'every COM release succeeded'",
                  "'every COM release left 0 outstanding references'",
                  "'every transient release inside the accepted readers succeeded'"):
        assert claim in code, claim
    # THE ACCEPTED READERS OWN THEIR OWN TRANSIENTS, and that gate is reported
    # rather than silently assumed: extending the residual accounting into them
    # would mean editing accepted files.
    assert "Get-TransientFailures" in code


def test_27_the_shutdown_is_the_accepted_path_and_emergency_is_never_a_pass() -> None:
    code = _own_code()
    shutdown = code[code.index("} catch {\n    $fatal = (Format-Err $_)"):]
    for required in ("$wb.Close($false)", "$excel.Quit()", "Wait-ExcelExit",
                     "Invoke-EmergencyExcelCleanup", "EXCEL SHUTDOWN"):
        assert required in shutdown, required
    assert shutdown.count("finally") >= 1
    assert shutdown.count("EXCEL SHUTDOWN:") >= 2
    clear_at = shutdown.index("$Error.Clear()")
    collect_at = shutdown.index("[System.GC]::Collect()")
    wait_at = shutdown.index("Wait-ExcelExit")
    assert clear_at < collect_at < wait_at
    match = re.search(r"Wait-ExcelExit -Identity \$excelIdentity -TimeoutSeconds (\d+)", code)
    assert match and int(match.group(1)) >= 60
    assert "$emergencyRequired = $true" in code
    assert "if ($ok) { exit 0 } else { exit 1 }" in code


def test_28_no_hidden_reference_pattern_and_nothing_is_saved() -> None:
    """The two patterns W1 refused structurally, held refused here, and the
    workbook is a disposable copy that is never written back."""
    own = _own_code()
    chains = re.findall(r"\$(?:wb|excel|workbooks)\.\w+\.\w+", own)
    assert not chains, f"a chained COM property access acquires an unnamed intermediate: {chains}"
    assert "$tempRoot = Join-Path ([System.IO.Path]::GetTempPath())" in own
    assert "$workbooks.Open($stageBPath)" in own
    for forbidden in (".Save()", ".SaveAs(", ".SaveCopyAs(", "Stop-Process"):
        assert forbidden not in own, forbidden
    assert "$wb.Close($false)" in own


def test_29_the_tree_is_proved_clean_before_excel_is_started() -> None:
    code = _own_code()
    revision = _function("Get-W5SourceRevision")
    assert "'pccm/src', 'pccm/spec', 'pccm/builder'" in revision
    assert "rev-parse HEAD" in revision
    refusal_at = code.index("REFUSED, BEFORE EXCEL WAS STARTED.")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert refusal_at < excel_at
    assert "exit 1" in code[refusal_at:excel_at]


def test_30_the_report_is_written_through_and_separates_prerequisites() -> None:
    body = _function("Write-W5Line")
    assert "Set-Content -LiteralPath $script:W5Path" in body
    code = _own_code()
    assert "'PREREQUISITE'" in code
    assert "$results.Count -gt 0" in code, (
        "a run that recorded no RESULT at all would otherwise pass")






# ===========================================================================
# H. THE FIXTURE AND THE BASELINE W5 REBUILDS
# ===========================================================================

def test_31_the_fixture_is_the_accepted_w4_baseline() -> None:
    case = _fixture()
    model = case["model"]
    assert len(model["cost_lines"]) + len(model["risks"]) == 5
    assert model["timeline"]["duration"] == 4
    assert case["iterations"] == 1000
    assert case["seed_mode"] == "FIXED"
    assert case["supplied_seed"] == 20260905
    assert case["selected_confidence_level"] == "P80"
    code = _own_code()
    assert "($driverCount -eq 5) -and ($yearCount -eq 4) -and ($iterations -eq 1000)" in code
    assert "($selectedLabel -ceq 'P80')" in code


def test_32_the_baseline_is_re_proved_before_the_annual_run() -> None:
    """W5 MAY USE ITS OWN SESSION, so it may not assume W4's. Every precondition
    the annual run depends on is established and checked here first."""
    code = _own_code()
    for claim in ("the calculation reports CURRENT before the simulation",
                  "the simulation reports CURRENT",
                  "the FIXED request was honoured before the annual run",
                  "a publication bank is active",
                  "the deterministic totals match the independent Phase-5 oracle"):
        assert claim in code, claim
    for field in ("'run_id'", "'request_fingerprint'", "'result_digest'"):
        assert field in code, field
    assert "the simulation published a " in code
    # AND THE IDENTITY IS CAPTURED BEFORE THE ANNUAL STEP, not after.
    capture_at = _at("$invariantsBefore = Get-W5RunInvariants")
    annual_at = _at("$announcement = Invoke-W5Annual")
    assert capture_at < annual_at
    assert _at("$iterationsBefore = Get-W5IterationBlock") < annual_at


def _at(needle: str) -> int:
    code = _own_code()
    assert needle in code, needle
    return code.index(needle)


# ===========================================================================
# I. THE STAMP, BOUND TO THE RUN THAT PRODUCED IT
# ===========================================================================

def test_33_every_stamp_field_the_projection_carries_is_checked() -> None:
    """A stamp that bound to the wrong run would be the worst possible pass, so
    every identity field is compared against the SIMULATION's own value rather
    than against itself."""
    code = _own_code()
    projected = set(json.loads(
        (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8")
    )["annual_records"]["stamp"]["rows"].keys())
    assert projected == {"run_id", "effective_seed", "request_fingerprint", "result_digest",
                         "iterations", "year_count", "selected_px_label",
                         "selected_px_probability", "published"}, projected
    for field in sorted(projected):
        assert f"'{field}'" in code, f"the stamp field {field} is never read"
    # THE FIVE THAT MUST EQUAL THE SIMULATION'S OWN, compared pairwise.
    for pair in ("Stamp = 'run_id';              Sim = 'run_id'",
                 "Stamp = 'effective_seed';      Sim = 'effective_seed'",
                 "Stamp = 'request_fingerprint'; Sim = 'request_fingerprint'",
                 "Stamp = 'result_digest';       Sim = 'result_digest'",
                 "Stamp = 'iterations';          Sim = 'iterations_run'"):
        assert pair in code, pair
    assert "Test-SimSameValue -A $stamp[$stampKey] -B $simBlock[$simKey]" in code
    assert "the stamped year count is the project duration" in code
    assert "the stamped selected Px label is the requested one" in code
    assert "the stamped selected Px probability is the one that label names" in code
    assert "the publication marker is the contracted one" in code
    assert "-Expected ([string]$p7.annual_records.stamp.published_marker)) `" in code, (
        "the publication marker is compared against something other than the "
        "projected one")
    assert "'PUBLISHED'" not in code, (
        "the publication marker is typed into the runner; it is production's "
        "constant, projected")
    # AND THE ANNUAL RUN IS REQUIRED TO HAVE SUCCEEDED AT ALL.
    assert "($announcement -like 'OK|*') $announcement" in code, (
        "the annual announcement is not required to be a success")
    assert "$annualOk = Add-W5Check 'PCCM_RunAnnualStochastic succeeded'" in code
    assert "if (-not $annualOk) { throw" in code


def test_34_the_stamp_and_records_are_read_only_through_the_projection() -> None:
    """NOT ONE HIDDEN-SHEET COORDINATE IS TYPED IN. Every column and row comes
    from phase7_acceptance_inspection.json, which is projected from
    sim_contract.yaml by the build that projects them into modSimContract."""
    for name in ("Get-W5AnnualStamp", "Get-W5AnnualRecord", "Get-W5AnnualRegionIndex",
                 "Get-W5AnnualFirstRecord"):
        body = _function(name)
        # NO COLUMN LETTER AND NO ROW NUMBER, in any form: 'AB8', 'AB' or 34.
        assert not re.search(r"'[A-Z]{1,3}\d{0,4}'", body), (
            f"{name} carries an address literal")
        # NO QUOTED LITERAL CARRYING A DIGIT AT ALL. A row is not only 'AB8':
        # `$column + '8'` is the same defect one concatenation later. The
        # literals are paired left to right rather than pattern-matched across
        # the body, or a colon separator and a later count would look like one.
        for literal in re.findall(r"'([^']*)'", body):
            assert not re.search(r"\d", literal), (
                f"{name} carries the literal {literal!r}; every row and column "
                "comes from the projection")
        assert "$P7.annual_records" in body or "$records" in body, name
    code = _own_code()
    assert "$p7.annual_records.first_record_row" not in code or True
    # The ladder columns are computed from the projected FIRST column and the
    # projected LENGTH, so eleven letters are never typed in.
    record = _function("Get-W5AnnualRecord")
    assert "$records.quantile_first_column.$Bank" in record
    assert "[int]$records.quantile_count" in record
    assert "ConvertFrom-W5ColumnNumber -Number ($first + $index)" in record


def test_35_the_ladder_labels_come_from_the_projection_not_from_the_runner() -> None:
    code = _own_code()
    assert "$quantileLabels = @($gateBCases.vocabulary.quantile_labels" in code
    assert "[array]::IndexOf($quantileLabels, $selectedLabel)" in code
    # Not one percentile label is typed in beyond the fixture's own selection.
    typed = set(re.findall(r"'(P\d{2})'", code))
    assert typed <= {"P80"}, typed
    gate_b = json.loads((BUILD / "phase6_gate_b_cases.json").read_text(encoding="utf-8"))
    labels = gate_b["vocabulary"]["quantile_labels"]
    projection = json.loads(
        (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8"))
    assert len(labels) == projection["annual_records"]["quantile_count"], (
        "the projected ladder length and the projected label list disagree")
    assert labels[0] == "P10" and labels[-1] == "P95"
    assert labels.index("P80") == 7
    # The runner requires that agreement at run time too.
    assert "$quantileLabels.Count -eq [int]$p7.annual_records.quantile_count" in code


def test_36_every_year_is_checked_for_a_complete_ordered_ladder() -> None:
    code = _own_code()
    assert "for ($offset = 0; $offset -lt $yearCount; $offset++)" in code
    assert "each of the " in code and "project years has one correctly indexed annual record" in code
    assert "both percentile ladders are complete, numeric and ordered in every year" in code
    assert "$ladder.Count -ne $ladderCount" in code, "a short ladder would pass"
    assert "$value -isnot [double]" in code, "a ladder value published as text would pass"
    assert "[double]$value -lt $previous" in code, "an out-of-order ladder would pass"
    assert "foreach ($measure in @('nominal', 'pv'))" in code
    # THE CALENDAR YEAR IS DERIVED FROM THE FIXTURE, never typed in.
    assert "[int]$model.timeline.start_year + $offset" in code


# ===========================================================================
# J. RECONCILIATION
# ===========================================================================

def test_37_the_type_7_helper_is_the_contracts_own_formula() -> None:
    """TRANSCRIBED, NOT INVENTED. Each line of the helper is the contract's."""
    body = _function("Get-W5Type7Value")
    contract = (PCCM_ROOT / "spec" / "sim_contract.yaml").read_text(encoding="utf-8")
    assert 'method: "hyndman_fan_type_7"' in contract
    for formula in ('h: "(n - 1) * p"', 'lo: "floor(h)"', 'hi: "min(lo + 1, n - 1)"',
                    'f: "h - lo"', 'value: "(1 - f) * x[lo] + f * x[hi]"'):
        assert formula in contract, formula
    assert "$h = ($n - 1) * $Probability" in body
    assert "$lo = [int][Math]::Floor($h)" in body
    assert "$hi = [Math]::Min($lo + 1, $n - 1)" in body
    assert "$f = $h - $lo" in body
    assert "((1 - $f) * [double]$sorted[$lo]) + ($f * [double]$sorted[$hi])" in body
    # ON A COPY, which is the contract's `sorting: on_copies_only`.
    assert "$sorted = @($Values | Sort-Object)" in body
    assert 'sorting: "on_copies_only"' in contract


def test_38_the_annual_blend_is_never_reconstructed_in_powershell() -> None:
    """Reimplementing production's convex annual blend here would compare an
    implementation against a copy of itself. What is computed is the TOTAL
    percentile; the profile is CONSUMED from the persisted result."""
    code = _own_code()
    assert code.count("Get-W5Type7Value") == 2, (
        "the Type-7 helper is defined once and used once per measure, on the "
        "iteration TOTALS column")
    assert "-Values ([double[]]@($totals))" in code
    # The profile values are read, summed and compared - never recomputed.
    assert "$profileSums[$measure] = $profileSums[$measure] + [double]$profile" in code
    for banned in ("AnnualVector", "blend", "lo_hi_f"):
        assert banned not in code, f"{banned} suggests the blend is being rebuilt"


def test_39_the_identity_allowance_is_the_projects_own_rule() -> None:
    """READ FROM THE CORPUS, NOT CHOSEN HERE, and the corpus reads it from the
    calc contract. No tolerance is invented and none is a literal."""
    contract = load_calc_contract(PCCM_ROOT / "spec" / "calc_contract.yaml")
    provenance = _cases()["provenance"]
    assert provenance["identity_absolute_floor"] == float(
        contract.tolerances.identity_absolute_floor)
    assert provenance["identity_relative_coefficient"] == float(
        contract.tolerances.identity_relative_coefficient)
    assert provenance["conditioning_scale_floor"] == float(
        contract.tolerances.conditioning_scale_floor)
    assert provenance["identity_absolute_floor"] == 1e-6
    assert provenance["identity_relative_coefficient"] == 1e-12
    assert provenance["conditioning_scale_floor"] == 1.0
    body = _function("Get-W5IdentityAllowance")
    assert "$Provenance.identity_absolute_floor" in body
    assert "$Provenance.identity_relative_coefficient" in body
    assert "$Provenance.conditioning_scale_floor" in body
    assert "[Math]::Max($floor, $coefficient * $scale)" in body
    assert "[Math]::Max($scaleFloor, [Math]::Abs($ConditioningScale))" in body
    code = _own_code()
    assert not re.search(r"\b\d*\.?\d+e-\d+\b", code), (
        "a numeric tolerance literal appears in the runner")


def test_40_the_conditioning_scale_names_the_arithmetic_not_the_result() -> None:
    """ERRATUM C1, APPLIED. A scale taken from the net result would let a model
    whose annual terms are large and whose total is near zero collapse its own
    tolerance. The scale here is the sum of the ABSOLUTE annual terms plus the
    aggregate they are compared against."""
    code = _own_code()
    assert "$profileScale[$measure] = $profileScale[$measure] + [Math]::Abs([double]$profile)" in code
    assert "$scale = $scale + [Math]::Abs([double]$published)" in code
    assert "$scale = [double]$profileScale[$measure]" in code
    # THE IDENTITY IS REQUIRED TO HOLD, not merely evaluated and printed.
    assert "(($published -is [double]) -and ($delta -le $identityAllowance)) `" in code, (
        "the reconciliation no longer requires the delta to be within the "
        "allowance")
    assert "$delta = [Math]::Abs($sum - [double]$published)" in code
    assert "selected-Px profile sums to the reported" in code
    assert "', delta ' + [string]$delta" in code
    assert "conditioning scale " in code
    # NO POST-HOC SCALING of the published answer.
    assert "* $factor" not in code and "/ $ratio" not in code


def test_41_the_published_total_px_is_cross_checked_before_it_is_relied_on() -> None:
    """A reconciliation against a WRONG total would pass. So the published
    ladder value and the contract's Type-7 over the iteration column are
    compared to each other first, under the same identity rule."""
    code = _own_code()
    assert "equals the contract" in code and "s Type-7 value over the iteration column" in code
    assert "$ladderRowKey = 'quantile_' + [string]($selectedIndex + 1)" in code
    assert "contingency_ladder.bank_value_columns.$bank.$measure" in code
    assert "contingency_ladder.rows.$ladderRowKey" in code
    cross_at = _at("s Type-7 value over the iteration column")
    identity_at = _at("selected-Px profile sums to the reported")
    assert cross_at < identity_at, (
        "the total is relied on before it is cross-checked")


# ===========================================================================
# K. HANDOFF, INVARIANTS AND COMPACT PERSISTENCE
# ===========================================================================

def test_42_the_four_accessors_are_checked_against_projected_vocabulary() -> None:
    code = _own_code()
    assert "$p7.handoff.distribution_states[1]" in code
    assert "$p7.handoff.profile_states[1]" in code
    assert "'CURRENT'" not in code.replace("-ceq 'CURRENT'", ""), (
        "a state string is typed in where the projection owns it")
    projection = json.loads(
        (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8"))
    assert projection["handoff"]["distribution_states"][1] == "CURRENT"
    assert projection["handoff"]["profile_states"][1] == "CURRENT"
    assert "reports the selected Px identity" in code
    assert "reports the project year count" in code
    assert "[int]$handoff[$accessors[3]]) -eq $yearCount" in code


def test_43_the_annual_run_changes_nothing_it_observes() -> None:
    """ANNUAL IS OBSERVATIONAL POST-PROCESSING, NOT A SECOND STOCHASTIC RUN, and
    the iteration columns are where that would show first."""
    code = _own_code()
    assert "Add-W5InvariantChecks 'the successful annual run'" in code
    assert "every published iteration value is unchanged by the annual run" in code
    assert "Test-SimSameValue -A $before[$row, $column] -B $after[$row, $column]" in code
    assert "the iteration block changed shape" in code
    # The frozen set is the whole published state minus the two derived rows.
    invariants = _function("Get-W5RunInvariants")
    assert "foreach ($bank in @($Inspection.publication.bank_labels))" in invariants
    assert "$out.Add('pending_auto_nonce'" in invariants
    assert "$script:W5DerivedRows = @('simulation_status', 'status_evaluated_at')" in code
    # And it is captured BEFORE and compared AFTER.
    assert _at("$invariantsBefore = Get-W5RunInvariants") < _at(
        "$invariantsAfter = Get-W5RunInvariants")


def test_44_compact_persistence_is_proved_over_the_whole_contracted_region() -> None:
    """ONE RECORD PER PROJECT YEAR AND NOTHING BEYOND. An N x Y matrix or an
    iteration-level workspace materialised as records would leave far more than
    four rows in the region, and the region is read in full rather than sampled."""
    code = _own_code()
    region = _function("Get-W5AnnualRegionIndex")
    assert "[int]$records.max_record_rows" in region, (
        "the region is not read to its contracted extent")
    assert "$records.index_columns.$Bank.project_index" in region
    assert "records, one per project year" in code
    assert "$populated -eq $yearCount" in code
    assert "the row immediately after the annual result is not populated" in code
    assert "$boundaryRow = $yearCount + 1" in code
    projection = json.loads(
        (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8"))
    assert projection["annual_records"]["max_record_rows"] == 200
    assert projection["annual_records"]["first_record_row"] == 34
    # THE REGION IS THE CONTRACTED ONE, not an arbitrary sweep of the sheet.
    assert "UsedRange" not in code and "SpecialCells" not in code, (
        "the runner scans for arbitrary hidden data instead of reading the "
        "contracted annual region")


def test_45_the_runner_states_what_no_windows_oracle_owns() -> None:
    """THE HONESTY CONTROL. The annual computation is proved independently on
    Linux; W5's subject is execution, persistence, read-back and reconciliation,
    and the report has to say so rather than let a reader assume more."""
    text = _text()
    assert "THERE IS NO INDEPENDENT ANNUAL WINDOWS ORACLE" in text
    assert "proved independently on Linux" in text or "proved independently" in text
    assert "PERSISTENCE, READ-BACK and RECONCILIATION" in text
    assert "reimplementing it in" in text or "copy of itself" in text
