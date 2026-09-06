#!/usr/bin/env python3
"""P7-7: the MINIMAL W7 runner, proved on Linux before a Windows session.

W7 PROVES TWO PERSISTENCE PROPERTIES AT ONCE, because one cannot be shown
without the other: the publication banks really cycle A -> B -> A, and the bank
that comes back round replaces a twenty-year annual answer with a four-year one
without leaving the tail of the old one readable.

THREE FIXED-SEED RUNS IN ONE SESSION. B changes the SEED, not the clock: a run
is told apart by its identity, and a different FIXED seed is a different
effective seed, a different request fingerprint and a different result digest -
reproducibly, on any machine, in any order.

THE AUTHORITY RULE IS THE CONTRACT'S. "A count is what says where the answer
stops - not the last non-blank row", so the runner reads the publication marker
plus the stamped year_count, and checks the physical clearing separately because
the stamp also declares that surplus rows are cleared. The whole former span is
read row by row - both index columns, every rung of both ladders, both profile
columns - never a scan of the sheet.

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
RUNNER = WINDOWS / "phase7_w7_bank_cycle.ps1"
W1 = WINDOWS / "phase7_w1_smoke.ps1"
W2 = WINDOWS / "phase7_w2_many_drivers.ps1"
W3 = WINDOWS / "phase7_w3_long_years.ps1"
W4 = WINDOWS / "phase7_w4_base_simulation.ps1"
W5 = WINDOWS / "phase7_w5_annual_success.ps1"
W6 = WINDOWS / "phase7_w6_selector_move.ps1"
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
# W7 writes the simulation request through it, so it is BOTH copied for the
# fixture's sake and called directly here.
CALLED_DIRECTLY = ("Set-NamedValue",)

# W1 THROUGH W6 ARE CLOSED. All three runners are accepted Windows evidence and
# are pinned, not maintained, by this round. The digests are LITERAL - hashing a
# file at import time would be a control that can never fail.
W1_SHA256 = "bf41a2193dbf2ea5fbddbe4dad3c6e94649f21262dcf9800ae71eda6e274b59f"
W2_SHA256 = "558bcfa0528b38c28c6e81dd726a18fed9078098f79bcabb792304e6255d8fee"
W3_SHA256 = "97af69bf836a2d798be4937f1e81cfe7218d2cfd702c1b96aa25623e7f8c12be"
W4_SHA256 = "9e1be0629d337b42b02bf790feef752c1fd37fbf0e6c400d4c90af3cd9d86cd0"
W5_SHA256 = "de2e3322c43e93da9647cca0004500e1427e47faf89cdc55c58c7b37e5159ea7"
W6_SHA256 = "feefceb0b7b0592a5799ffd5f95c9ac02c1d7056dfaf04d8d20bf868e7e7f019"

_CACHE: dict = {}


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    """The runner with its block comment and line comments removed.

    Its header explains what W7 refuses to do and names the things the controls
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
    assert body is not None, f"{name} is not defined in the W7 runner"
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
    """THE W7 CASE. It is the only one that carries two durations and two seeds,
    which is what a bank cycle with a shrink needs."""
    case = [s for s in _cases()["scenarios"] if s["id"] == "W7"]
    assert case, "the acceptance corpus carries no W7 scenario"
    return case[0]


def _year_window() -> tuple[int, int, int]:
    """(min_year, max_year, max_generated_year_columns) from the structural contract.

    THE AUTHORITY, NOT A REMEMBERED NUMBER. The W7 fixture's calendar span is
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
    """(first calendar year, last calendar year, duration) of the W7 fixture."""
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
# A. SCOPE: THIS IS W7, AND W1 IS CLOSED
# ===========================================================================

def test_01_the_runner_exists_and_is_dedicated() -> None:
    assert RUNNER.exists()
    lines = _text().splitlines()
    assert len(lines) < 1800, (
        f"{len(lines)} lines; W7 rebuilds a baseline and then proves two moves "
        "on one simulation, and most of its behaviour is meant to come from the "
        "accepted Phase-5 fixture, the accepted Phase-6 state readers and the "
        "projection")
    # Most of the file is either the accepted copied helpers or W7's own small
    # comparison; neither should have grown into a second harness.
    assert len(_own_code().splitlines()) < 1400


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
    """W7 runs the annual endpoint ONCE, expecting a refusal, and one simulation.

    Everything else stochastic belongs to another scenario, and a second
    stochastic oracle invented here would be exactly the "compare a run against
    another run and call it independence" the authorisation refuses.
    """
    code = _own_code()
    for forbidden in ("PCCM_RunSensitivity", "Invoke-Phase5GateBScenarios",
                      "Invoke-Phase6GateBScenarios", "phase7_acceptance_scenarios",
                      "phase7_w1_smoke", "phase7_w2_many_drivers", "phase7_w3_long_years"):
        assert forbidden not in code, f"{forbidden} is outside W7"
    for scenario in ("'W2'", "'W3'", "'W5'", "'W6'", "'W8'"):
        assert scenario not in code, scenario
    assert "$_.id -ceq 'W7'" in code
    # W7 REUSES THE W4 BEHAVIOURAL FIXTURE DELIBERATELY: the annual answer it
    # W7 SELECTS ITS OWN CORPUS CASE - the one that carries both durations and
    # both seeds.
    assert "the W7 fixture shrinks and the two seeds differ" in code
    # THREE RUNS, ONE WORKFLOW. The simulation and the annual endpoint each have
    # exactly one call site, inside the helper the three runs share, so a fourth
    # run cannot appear without the helper being called again.
    assert code.count("Invoke-Phase6Simulation") == 1, (
        "the simulation is invoked from more than one place; the three runs go "
        "through one workflow helper")
    assert code.count("Invoke-W7Annual") == 2, (
        "the annual endpoint is defined once and invoked from one place")
    assert code.count("Invoke-W7Run -Excel $excel") == 3, (
        "W7 is three runs: A1, B and A2")


def test_05_it_does_not_repeat_the_w1_surface_matrix() -> None:
    """One lightweight read-only compile trigger is authorised because Excel
    must execute the current project. The surface matrix is W1's."""
    code = _own_code()
    for forbidden in ("VBProject", "VBComponents", "CodeModule", "ProcOfLine"):
        assert forbidden not in code, f"{forbidden} is W1's evidence, not W7's"
    assert "$excel.Run('PCCM_CalculationStatus')" in code
    assert "the current VBAProject compiles in real Excel" in code


def test_06_the_accepted_w1_and_w2_runners_are_untouched() -> None:
    """BOTH ARE CLOSED on accepted Windows evidence. This round does not modify
    or rerun either, and an edit to one must fail here rather than pass."""
    assert all(path.exists() for path in (W1, W2, W3, W4, W5, W6))
    assert hashlib.sha256(W1.read_bytes()).hexdigest() == W1_SHA256, (
        "the accepted W1 runner has been modified; it is closed evidence")
    assert hashlib.sha256(W2.read_bytes()).hexdigest() == W2_SHA256, (
        "the accepted W2 runner has been modified; it is closed evidence")
    assert hashlib.sha256(W3.read_bytes()).hexdigest() == W3_SHA256, (
        "the accepted W3 runner has been modified; it is closed evidence")
    assert hashlib.sha256(W4.read_bytes()).hexdigest() == W4_SHA256, (
        "the accepted W4 runner has been modified; it is closed evidence")
    assert hashlib.sha256(W5.read_bytes()).hexdigest() == W5_SHA256, (
        "the accepted W5 runner has been modified; it is closed evidence")
    assert hashlib.sha256(W6.read_bytes()).hexdigest() == W6_SHA256, (
        "the accepted W6 runner has been modified; it is closed evidence")
    code = _code()
    for closed in ("phase7_w1_smoke", "phase7_w2_many_drivers",
                   "phase7_w3_long_years", "phase7_w4_base_simulation",
                   "phase7_w5_annual_success", "phase7_w6_selector_move"):
        assert closed not in code, f"W7 reaches into the closed {closed} at run time"


# THE PARTS W7 SHARES WITH W2, PINNED SO THE MANDATED COPY CANNOT DRIFT.
# Reusing W2's proven execution architecture was the instruction; two runners
# that started identical and quietly diverged would be the cost of it, so the
# shared functions are compared after normalising the scenario name.
SHARED_WITH_W2 = ("Write-W7Line", "Add-W7Check", "Format-W7Value", "Invoke-W7Release",
                  "Get-W7SourceRevision", "Compare-W7Cell", "Compare-W7Table")


def test_06b_the_shape_shared_with_w2_is_identical_to_w2s() -> None:
    import re as _re
    ours = _text()
    theirs = W2.read_text(encoding="utf-8")
    for name in SHARED_WITH_W2:
        mine = _re.search(rf"^function {name} \{{(.*?)^\}}", ours, _re.M | _re.S)
        assert mine, f"{name} is not defined in the W7 runner"
        w2name = name.replace("W7", "W2")
        yours = _re.search(rf"^function {w2name} \{{(.*?)^\}}", theirs, _re.M | _re.S)
        assert yours, f"{w2name} is no longer in the accepted W2 runner"
        assert mine.group(1).replace("W7", "W2") == yours.group(1), (
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
        assert ours, f"{name} is not defined in the W7 runner"
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
    # `Set-W7NamedText` writes one defined name: its three objects are acquired
    # and released inside a single statement, in the same convention every
    # accepted reader uses, and they are covered by the accepted-reader
    # transient gate. Everything the runner HOLDS goes through its own ledger.
    outside = code.replace(_function("Set-W7NamedText"), "")
    assert "Release-Transient" not in outside, (
        "a release the runner holds still goes through the silent-on-success "
        "helper; only Set-W7NamedText may use it")
    assert _function("Set-W7NamedText").count("Release-Transient") == 3
    assert "Invoke-NamedRelease" not in code, (
        "Invoke-NamedRelease keeps the release count to itself")
    for label in ("'Workbook'", "'Workbooks'", "'Excel.Application'"):
        assert f"Invoke-W7Release $rel" in code and label in code, label
    ledger_at = code.index("$rel = New-ReleaseLedger")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert ledger_at < excel_at
    assert code.count("$comAcquired = $comAcquired + 1") == 3, (
        "the three session objects are each counted where they are acquired")
    # AND EVERY RANGE READ IS COUNTED TOO. W7 opens three of them - the
    # iteration column before and after the annual run, and the annual region -
    # and an uncounted acquisition would break the balance check without
    # anything saying which one.
    # W7 opens FOUR range reads: the iteration column at the baseline, after the
    # selector move and after the rerun, and the annual region at the end.
    # W7 opens FOUR iteration reads: one per run, plus bank B re-read after A2
    # committed, which is how bank isolation is proved value for value.
    counted = re.findall(r"\$comAcquired = \$comAcquired \+ \[int\]\$(\w+)\.Acquired", code)
    assert sorted(counted) == ["bIterationsAfter", "iterationsA1", "iterationsA2",
                               "iterationsB"], counted
    for name in ("Get-W7IterationBlock", "Get-W7AnnualRegionIndex"):
        block = _function(name)
        assert block.count("Invoke-W7Release $Ledger") == 3, name
        assert block.count("$acquired = $acquired + 1") == 3, name
    helper = _function("Invoke-W7Release")
    assert "Release-ComObjectSafe" in helper
    assert "[int]$rec.Count -ne 0" in helper
    assert "$script:W7Residual.Add" in helper


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
    revision = _function("Get-W7SourceRevision")
    assert "'pccm/src', 'pccm/spec', 'pccm/builder'" in revision
    assert "rev-parse HEAD" in revision
    refusal_at = code.index("REFUSED, BEFORE EXCEL WAS STARTED.")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert refusal_at < excel_at
    assert "exit 1" in code[refusal_at:excel_at]


def test_30_the_report_is_written_through_and_separates_prerequisites() -> None:
    body = _function("Write-W7Line")
    assert "Set-Content -LiteralPath $script:W7Path" in body
    code = _own_code()
    assert "'PREREQUISITE'" in code
    assert "$results.Count -gt 0" in code, (
        "a run that recorded no RESULT at all would otherwise pass")










# ===========================================================================
# H. THE THREE FIXTURES AND THE BANK CYCLE
# ===========================================================================

def _at(needle: str) -> int:
    code = _own_code()
    assert needle in code, needle
    return code.index(needle)


def _publication() -> dict:
    return json.loads(
        (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8")
    )["publication_semantics"]


def test_31_the_corpus_carries_both_durations_and_two_distinct_seeds() -> None:
    case = _fixture()
    long_model, short_model = case["model"], case["shrink_model"]
    assert long_model["timeline"]["duration"] == 20
    assert short_model["timeline"]["duration"] == 4
    # ONLY THE DURATION SHRINKS: the same drivers, so a changed row count cannot
    # be explained by a changed model.
    assert len(short_model["cost_lines"]) == len(long_model["cost_lines"])
    assert len(short_model["risks"]) == len(long_model["risks"])
    assert len(long_model["cost_lines"]) + len(long_model["risks"]) == 5
    assert case["seed_mode"] == "FIXED"
    assert case["supplied_seed"] == 20260906
    assert case["second_supplied_seed"] == 20260907
    assert case["supplied_seed"] != case["second_supplied_seed"]
    assert case["iterations"] == 1000
    assert case["selected_confidence_level"] == "P80"
    # THE SHRINK MODEL HAS ITS OWN ORACLE EXPECTATION, so a later scenario can
    # compare it without recomputing anything.
    assert len(case["shrink_expected"]["calc_years"]) == 4
    assert len(case["expected"]["calc_years"]) == 20


def test_32_b_is_a_different_request_not_a_different_moment() -> None:
    """A run is told apart by its identity. B changes the SEED, so its effective
    seed, request fingerprint and result digest differ reproducibly - on any
    machine, in any order, with no dependence on a clock."""
    code = _own_code()
    assert "$secondSeed = [double]$case.second_supplied_seed" in code
    assert "-Seed $secondSeed" in code
    assert "the B run is a different request from A1" in code
    assert "'effective_seed'" in code and "'request_fingerprint'" in code
    for banned in ("Get-Date -Format", "[DateTime]::Now", "Get-Random"):
        assert banned not in code, banned
    # And B does NOT change the model, so no recalculation is needed: only the
    # request moves.
    assert "-Model $null -Seed $secondSeed" in code


def test_33_the_bank_cycle_comes_from_the_projection() -> None:
    """The runner never assumes a blank selector means A, or that A is followed
    by B: it asks the projected candidate-target map."""
    publication = _publication()
    targets = {entry["active_bank"]: entry["candidate_bank"]
               for entry in publication["candidate_target"]}
    assert targets == {None: "A", "A": "B", "B": "A"}
    assert publication["bank_labels"] == ["A", "B"]
    contract = (PCCM_ROOT / "spec" / "sim_contract.yaml").read_text(encoding="utf-8")
    assert re.search(r'candidate_target:\s*\n\s*"":\s*"A"\s*\n\s*"A":\s*"B"\s*\n\s*"B":\s*"A"',
                     contract), "the contract's bank cycle is no longer A then B then A"
    code = _own_code()
    body = _function("Get-W7CandidateBank")
    assert "$P7.publication_semantics.candidate_target" in body
    assert code.count("Get-W7CandidateBank -P7 $p7") >= 6, (
        "the expected bank is asked for at each step and in the preflight")
    assert "$expectedFirstBank = Get-W7CandidateBank" in code
    assert "$expectedSecondBank = Get-W7CandidateBank -P7 $p7 -ActiveBank $bankA" in code
    assert "$expectedThirdBank = Get-W7CandidateBank -P7 $p7 -ActiveBank $bankB" in code
    assert "the projected bank cycle is two banks and three transitions" in code
    # NO DIRECT WRITE TO BANK-SELECTION STATE.
    for banned in ("Set-SimField", "Set-SimRawCell", "active_bank'  -Value", "Set-SimPending"):
        assert banned not in code, banned


def test_34_the_run_id_progression_is_required_not_observed() -> None:
    code = _own_code()
    assert "the B run id advanced from A1" in code
    assert "the A2 run id advanced from B" in code
    assert "-Expected ([double]$runIdA1 + 1)" in code
    assert "-Expected ([double]$runIdB + 1)" in code
    assert "the A2 identity is distinct from both A1 and B" in code
    contract = (PCCM_ROOT / "spec" / "sim_contract.yaml").read_text(encoding="utf-8")
    assert "last_run_id + 1" in contract, (
        "the contract's run-id allocation is no longer last + 1")


# ===========================================================================
# I. THE AUTHORITY RULE
# ===========================================================================

def test_35_authority_is_the_marker_plus_the_year_count() -> None:
    """"A count is what says where the answer stops - not the last non-blank
    row." The runner reads authority the contract's way, and a control forbids
    the shortcut."""
    publication = _publication()
    assert publication["published_written_last"] is True
    assert publication["cleared_before_write"] is True
    assert publication["surplus_rows_cleared"] is True
    assert "the last non-blank row is never the authority" in publication["authority_rule"]
    contract = (PCCM_ROOT / "spec" / "sim_contract.yaml").read_text(encoding="utf-8")
    for line in ("published_written_last: true", "cleared_before_write: true",
                 "surplus_rows_cleared: true"):
        assert line in contract, line
    # PROJECTED FROM THE CONTRACT, NOT RETYPED IN THE BUILDER.
    generator = (PCCM_ROOT / "builder" / "pccm_builder"
                 / "phase7_acceptance.py").read_text(encoding="utf-8")
    for key in ("published_written_last", "cleared_before_write", "surplus_rows_cleared"):
        assert f'bool(stamp["{key}"])' in generator, key
    assert 'sim.raw["publication"]["banks"]["candidate_target"]' in generator
    assert 'sim.raw["publication"]["banks"]["labels"]' in generator
    body = _function("Test-W7Authoritative")
    assert "$P7.annual_records.stamp.published_marker" in body, (
        "the publication marker is typed in rather than projected")
    assert "-Actual $Stamp['year_count']" in body
    assert "$markerOk -and $countOk" in body, (
        "authority is decided by one fact where the contract names two")
    code = _own_code()
    assert code.count("Test-W7Authoritative -Stamp") == 3, (
        "each of the three annual results is checked for authority")
    assert "UsedRange" not in code and "SpecialCells" not in code


def test_36_the_builder_refuses_a_contract_that_stopped_clearing() -> None:
    """If surplus rows were no longer cleared, "row 5 is blank" would be a
    property no scenario could require, so the emitter refuses it."""
    from pccm_builder import phase7_acceptance as emitter

    spec = PCCM_ROOT / "spec"
    sim = load_sim_contract(spec / "sim_contract.yaml")
    limits = load_structure_contract(spec / "structure_contract.yaml").limits
    inspection = emitter.build_phase7_inspection(sim, limits.max_generated_year_columns)
    cases = emitter.build_phase7_cases(
        sim=sim, calc=load_calc_contract(spec / "calc_contract.yaml"),
        max_record_rows=limits.max_generated_year_columns,
        min_year=limits.min_year, max_year=limits.max_year)
    emitter.validate_phase7_artifacts(inspection, cases)
    broken = json.loads(json.dumps(inspection))
    broken["publication_semantics"]["surplus_rows_cleared"] = False
    with pytest.raises(ValueError, match="surplus"):
        emitter.validate_phase7_artifacts(broken, cases)
    cycled = json.loads(json.dumps(inspection))
    cycled["publication_semantics"]["bank_labels"] = ["A", "B", "C"]
    with pytest.raises(ValueError, match="bank cycle"):
        emitter.validate_phase7_artifacts(cycled, cases)


# ===========================================================================
# J. CAPTURE, ISOLATION AND THE SHRINK
# ===========================================================================

def test_37_a_bank_capture_is_the_whole_bank_as_plain_data() -> None:
    body = _function("Get-W7BankCapture")
    assert "Get-SimBankBlock" in body, "the simulation block is not captured"
    assert "Get-W7AnnualSurface" in body, "the annual answer is not captured"
    assert "'sim.'" in body and "'annual.'" in body
    surface = _function("Get-W7AnnualSurface")
    assert "Get-W7AnnualStamp" in surface and "Get-W7AnnualRecord" in surface
    assert "[int]$P7.annual_records.quantile_count" in surface
    # THE LADDERS ARE FLATTENED RUNG BY RUNG in the capture as well as in the
    # residue reader. A capture that held each ladder as one array would compare
    # arrays rather than values, and a changed rung could pass.
    assert "$name.StartsWith('ladder_')" in surface, (
        "the capture does not flatten the ladders; a changed rung could pass")
    assert "for ($index = 0; $index -lt $ladderCount; $index++)" in surface
    compare = _function("Compare-W7Surface")
    assert "Test-SimSameValue" in compare, (
        "the capture comparison does not compare types as well as values")


def test_38_each_inactive_bank_is_proved_untouched() -> None:
    code = _own_code()
    assert "is value-identical after B committed" in code
    assert "is value-identical after A2 committed" in code
    assert "Compare-W7Surface -Before $captureA1 -After $captureA1AfterB" in code
    assert "Compare-W7Surface -Before $captureB -After $captureBAfterA2" in code
    assert "changed when A2 committed" in code
    assert "Compare-W7IterationGrid -Before $iterationsB.Values" in code
    # The captures are taken BEFORE the run that must not disturb them.
    assert _at("$captureA1 = Get-W7BankCapture") < _at("B - THE NEXT RUN")
    assert _at("$captureB = Get-W7BankCapture") < _at("A2 - BACK TO BANK")


def test_39_the_shrink_is_checked_across_the_whole_former_span() -> None:
    """Not row 5 alone: every row A1 used to occupy beyond the new count, in
    both index columns, every rung of both ladders and both profile columns."""
    code = _own_code()
    assert "-FromRow ($shortYears + 1) -ToRow $longYears" in code, (
        "the residue check does not span the whole former authoritative range")
    assert "no A1 record survives anywhere in former rows " in code
    body = _function("Get-W7SurplusResidue")
    assert "for ($offset = $FromRow - 1; $offset -le $ToRow - 1; $offset++)" in body
    assert "Get-W7AnnualRecord" in body, (
        "the residue reader does not read the contracted record columns")
    assert "$name.StartsWith('ladder_')" in body, "the ladders are not scanned"
    assert "Test-SimBlank" in body
    # AND ROWS 1..4 CARRY A2, so the shrink is a replacement rather than a wipe.
    assert "carry the A2 answer" in code
    assert "A2 is authoritative for exactly " in code
    # The A1 tail was already blank before the shrink, so the check measures a
    # change rather than a state that was always true.
    assert "-FromRow ($longYears + 1) -ToRow ($longYears + 1)" in code
    assert "s answer is blank" in code


def test_40_the_a2_stamp_belongs_to_the_a2_run() -> None:
    code = _own_code()
    assert "the A2 annual stamp belongs to the A2 run" in code
    assert "Test-SimSameValue -A $stampA2['run_id'] -B $runIdA2" in code
    assert "$stampA2['result_digest']" in code
    assert "the A2 applied timeline is the shorter one" in code
    assert "A2 cycled back to bank " in code, (
        "nothing requires the third run to return to the first bank")
    assert "($bankA2 -ceq $expectedThirdBank) -and ($bankA2 -ceq $bankA)" in code, (
        "the return to bank A is not required to be BOTH the projected target "
        "and the bank A1 used")
    assert "A2 reports both annual products CURRENT and the short year count" in code
    assert "the A2 profile is stamped " in code


# ===========================================================================
# K. RECONCILIATION, NONCE AND WORKFLOW DISCIPLINE
# ===========================================================================

def test_41_all_three_annual_results_are_reconciled_by_one_function() -> None:
    code = _own_code()
    assert code.count("Invoke-W7Reconciliation") == 4, (
        "the reconciliation is defined once and invoked once per annual result")
    for stage in ("-Stage 'A1'", "-Stage 'B'", "-Stage 'A2'"):
        assert stage in code, stage
    body = _function("Invoke-W7Reconciliation")
    assert "TOTAL equals the contract" in body
    assert "selected-Px profile sums to the reported" in body
    assert "contingency is " in body
    assert "$semantics.contingency_formula" in body
    assert "$semantics.baseline_metric_key" in body
    assert "Get-W7Type7Value" in body
    assert "Get-W7IdentityAllowance -Provenance $Provenance" in body
    # EACH RUN'S OWN ITERATION COLUMN, never another run's.
    assert "-Grid $iterationsA1.Values -Bank $bankA " in code
    assert "-Grid $iterationsB.Values -Bank $bankB " in code
    assert "-Grid $iterationsA2.Values -Bank $bankA2 " in code
    assert not re.search(r"\b\d*\.?\d+e-\d+\b", code), (
        "a numeric tolerance literal appears in the runner")


def test_42_the_fixed_seed_nonce_discipline_is_proved_across_all_three_runs() -> None:
    code = _own_code()
    assert "no FIXED run consumed an AUTO nonce" in code
    assert "the AUTO nonce and its pending marker never moved" in code
    assert "$nonceStart = $stateStart['shared']['next_auto_nonce']" in code
    assert "$pendingStart = $stateStart['pending_auto_nonce']" in code
    # Captured before the first run and compared after the last.
    assert _at("$nonceStart =") < _at("A1 - ")
    assert code.count("'consumed_auto_nonce'") >= 1
    # W7 stays FIXED throughout: no AUTO behaviour is introduced.
    assert "'AUTO'" not in code


def test_43_the_workflow_is_one_helper_and_recalculates_only_when_the_model_moves() -> None:
    """B changes only the request, so no recalculation is needed; A1 and A2
    change the model, so both recalculate through production's own endpoint."""
    body = _function("Invoke-W7Run")
    assert "Set-Phase5Fixture" in body
    assert "-Operation 'PCCM_Calculate'" in body
    assert "if ($Recalculate) {" in body
    assert "Invoke-Phase6Simulation" in body
    assert "Invoke-W7Annual" in body
    code = _own_code()
    assert code.count("-Label 'A1' -Recalculate") == 1
    assert code.count("-Label 'A2' -Recalculate") == 1
    assert "-Label 'B'" in code
    assert "-Iterations $iterations -Label 'B')" in code, (
        "the B run recalculates, which would mean its model moved")
    # THE MEASURES AND THE SELECTOR STILL COME FROM THE PROJECTION.
    assert "$measures = @($p7.summary_semantics.contingency_measures" in code
    assert "@('nominal', 'pv')" not in code
    assert "$inspection.inputs.($p7.selector_semantics.selector_input_key).defined_name" in code
    assert not re.search(r"'inp\w+'", code)


def test_44_the_runner_states_what_no_windows_oracle_owns() -> None:
    text = _text()
    assert "THERE IS NO INDEPENDENT ANNUAL WINDOWS ORACLE" in text
    assert "not only a storage shape" in text or "not only a storage" in text
