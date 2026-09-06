#!/usr/bin/env python3
"""P7-7: the MINIMAL W6 runner, proved on Linux before a Windows session.

W6 PROVES A DIFFERENCE, not a value. sim_contract.yaml says the two annual
products do not share a currentness rule: the per-year percentile LADDERS are a
property of the run alone and no selector enters them, while the selected-Px
PROFILE is the blend at ONE resolved Px and stops being current the moment the
selector resolves to a different one - and is never relabelled. Those four
booleans are projected now, so the runner states the rule it tests rather than
encoding it, and a control requires it to assert them before testing anything.

TWO MOVES, ONE SIMULATION. The selector moves P80 to P50 and NOTHING else runs;
then the annual endpoint alone reruns. Across both, the run identity, the
fingerprint, the digest, the nonce, the pending marker and every published
iteration value must be unchanged, and the distribution rungs must be
value-identical - a reporting selector may not recompute a distribution. What
must move is the profile and its stamp.

ONE THING IS DELIBERATELY NOT A FAILURE, and a control holds it that way: the
P50 profile is expected to differ from the P80 one, but a fixture COULD produce
equal values, and failing on equality would assert a property of this fixture
rather than of the code. The change is reported; correctness is proved by the
reconciliation and by the stamp, which must move.

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
RUNNER = WINDOWS / "phase7_w6_selector_move.ps1"
W1 = WINDOWS / "phase7_w1_smoke.ps1"
W2 = WINDOWS / "phase7_w2_many_drivers.ps1"
W3 = WINDOWS / "phase7_w3_long_years.ps1"
W4 = WINDOWS / "phase7_w4_base_simulation.ps1"
W5 = WINDOWS / "phase7_w5_annual_success.ps1"
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
# W6 writes the simulation request through it, so it is BOTH copied for the
# fixture's sake and called directly here.
CALLED_DIRECTLY = ("Set-NamedValue",)

# W1 THROUGH W5 ARE CLOSED. All three runners are accepted Windows evidence and
# are pinned, not maintained, by this round. The digests are LITERAL - hashing a
# file at import time would be a control that can never fail.
W1_SHA256 = "bf41a2193dbf2ea5fbddbe4dad3c6e94649f21262dcf9800ae71eda6e274b59f"
W2_SHA256 = "558bcfa0528b38c28c6e81dd726a18fed9078098f79bcabb792304e6255d8fee"
W3_SHA256 = "97af69bf836a2d798be4937f1e81cfe7218d2cfd702c1b96aa25623e7f8c12be"
W4_SHA256 = "9e1be0629d337b42b02bf790feef752c1fd37fbf0e6c400d4c90af3cd9d86cd0"
W5_SHA256 = "de2e3322c43e93da9647cca0004500e1427e47faf89cdc55c58c7b37e5159ea7"

_CACHE: dict = {}


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    """The runner with its block comment and line comments removed.

    Its header explains what W6 refuses to do and names the things the controls
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
    assert body is not None, f"{name} is not defined in the W6 runner"
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
    """THE W4 BEHAVIOURAL CASE. W6 reuses it deliberately: the annual answer it
    proves has to be the answer produced from the baseline W4 accepted."""
    case = [s for s in _cases()["scenarios"] if s["id"] == "W4"]
    assert case, "the acceptance corpus carries no W4 behavioural fixture"
    return case[0]


def _year_window() -> tuple[int, int, int]:
    """(min_year, max_year, max_generated_year_columns) from the structural contract.

    THE AUTHORITY, NOT A REMEMBERED NUMBER. The W6 fixture's calendar span is
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
    """(first calendar year, last calendar year, duration) of the W6 fixture."""
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
# A. SCOPE: THIS IS W6, AND W1 IS CLOSED
# ===========================================================================

def test_01_the_runner_exists_and_is_dedicated() -> None:
    assert RUNNER.exists()
    lines = _text().splitlines()
    assert len(lines) < 1800, (
        f"{len(lines)} lines; W6 rebuilds a baseline and then proves two moves "
        "on one simulation, and most of its behaviour is meant to come from the "
        "accepted Phase-5 fixture, the accepted Phase-6 state readers and the "
        "projection")
    # Most of the file is either the accepted copied helpers or W6's own small
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
    """W6 runs the annual endpoint ONCE, expecting a refusal, and one simulation.

    Everything else stochastic belongs to another scenario, and a second
    stochastic oracle invented here would be exactly the "compare a run against
    another run and call it independence" the authorisation refuses.
    """
    code = _own_code()
    for forbidden in ("PCCM_RunSensitivity", "Invoke-Phase5GateBScenarios",
                      "Invoke-Phase6GateBScenarios", "phase7_acceptance_scenarios",
                      "phase7_w1_smoke", "phase7_w2_many_drivers", "phase7_w3_long_years"):
        assert forbidden not in code, f"{forbidden} is outside W6"
    for scenario in ("'W2'", "'W3'", "'W5'", "'W7'", "'W8'"):
        assert scenario not in code, scenario
    # W6 REUSES THE W4 BEHAVIOURAL FIXTURE DELIBERATELY: the annual answer it
    # proves has to be the answer produced from the baseline W4 accepted, so the
    # corpus case it selects is W4's and the runner says so.
    assert "$_.id -ceq 'W4'" in code
    assert "the W6 fixture is the accepted W5 behavioural baseline" in code
    # ONE simulation, ONE annual attempt.
    assert code.count("Invoke-Phase6Simulation") == 1, (
        "W6 runs the simulation more than once; a same-seed replay compared "
        "against itself is not independent evidence and is not authorised here")
    # W6 RUNS THE ANNUAL ENDPOINT TWICE ON PURPOSE - once for the baseline and
    # once after the selector moves - and that is the whole scenario.
    assert code.count("Invoke-W6Annual") == 3, (
        "the annual endpoint is defined once and invoked exactly twice")


def test_05_it_does_not_repeat_the_w1_surface_matrix() -> None:
    """One lightweight read-only compile trigger is authorised because Excel
    must execute the current project. The surface matrix is W1's."""
    code = _own_code()
    for forbidden in ("VBProject", "VBComponents", "CodeModule", "ProcOfLine"):
        assert forbidden not in code, f"{forbidden} is W1's evidence, not W6's"
    assert "$excel.Run('PCCM_CalculationStatus')" in code
    assert "the current VBAProject compiles in real Excel" in code


def test_06_the_accepted_w1_and_w2_runners_are_untouched() -> None:
    """BOTH ARE CLOSED on accepted Windows evidence. This round does not modify
    or rerun either, and an edit to one must fail here rather than pass."""
    assert W1.exists() and W2.exists() and W3.exists() and W4.exists() and W5.exists()
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
    code = _code()
    for closed in ("phase7_w1_smoke", "phase7_w2_many_drivers",
                   "phase7_w3_long_years", "phase7_w4_base_simulation",
                   "phase7_w5_annual_success"):
        assert closed not in code, f"W6 reaches into the closed {closed} at run time"


# THE PARTS W6 SHARES WITH W2, PINNED SO THE MANDATED COPY CANNOT DRIFT.
# Reusing W2's proven execution architecture was the instruction; two runners
# that started identical and quietly diverged would be the cost of it, so the
# shared functions are compared after normalising the scenario name.
SHARED_WITH_W2 = ("Write-W6Line", "Add-W6Check", "Format-W6Value", "Invoke-W6Release",
                  "Get-W6SourceRevision", "Compare-W6Cell", "Compare-W6Table")


def test_06b_the_shape_shared_with_w2_is_identical_to_w2s() -> None:
    import re as _re
    ours = _text()
    theirs = W2.read_text(encoding="utf-8")
    for name in SHARED_WITH_W2:
        mine = _re.search(rf"^function {name} \{{(.*?)^\}}", ours, _re.M | _re.S)
        assert mine, f"{name} is not defined in the W6 runner"
        w2name = name.replace("W6", "W2")
        yours = _re.search(rf"^function {w2name} \{{(.*?)^\}}", theirs, _re.M | _re.S)
        assert yours, f"{w2name} is no longer in the accepted W2 runner"
        assert mine.group(1).replace("W6", "W2") == yours.group(1), (
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
        assert ours, f"{name} is not defined in the W6 runner"
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
    # `Set-W6NamedText` writes one defined name: its three objects are acquired
    # and released inside a single statement, in the same convention every
    # accepted reader uses, and they are covered by the accepted-reader
    # transient gate. Everything the runner HOLDS goes through its own ledger.
    outside = code.replace(_function("Set-W6NamedText"), "")
    assert "Release-Transient" not in outside, (
        "a release the runner holds still goes through the silent-on-success "
        "helper; only Set-W6NamedText may use it")
    assert _function("Set-W6NamedText").count("Release-Transient") == 3
    assert "Invoke-NamedRelease" not in code, (
        "Invoke-NamedRelease keeps the release count to itself")
    for label in ("'Workbook'", "'Workbooks'", "'Excel.Application'"):
        assert f"Invoke-W6Release $rel" in code and label in code, label
    ledger_at = code.index("$rel = New-ReleaseLedger")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert ledger_at < excel_at
    assert code.count("$comAcquired = $comAcquired + 1") == 3, (
        "the three session objects are each counted where they are acquired")
    # AND EVERY RANGE READ IS COUNTED TOO. W6 opens three of them - the
    # iteration column before and after the annual run, and the annual region -
    # and an uncounted acquisition would break the balance check without
    # anything saying which one.
    # W6 opens FOUR range reads: the iteration column at the baseline, after the
    # selector move and after the rerun, and the annual region at the end.
    counted = re.findall(r"\$comAcquired = \$comAcquired \+ \[int\]\$(\w+)\.Acquired", code)
    assert sorted(counted) == ["iterationsAfterMove", "iterationsAfterRerun",
                               "iterationsBefore", "region"], counted
    for name in ("Get-W6IterationBlock", "Get-W6AnnualRegionIndex"):
        block = _function(name)
        assert block.count("Invoke-W6Release $Ledger") == 3, name
        assert block.count("$acquired = $acquired + 1") == 3, name
    helper = _function("Invoke-W6Release")
    assert "Release-ComObjectSafe" in helper
    assert "[int]$rec.Count -ne 0" in helper
    assert "$script:W6Residual.Add" in helper


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
    revision = _function("Get-W6SourceRevision")
    assert "'pccm/src', 'pccm/spec', 'pccm/builder'" in revision
    assert "rev-parse HEAD" in revision
    refusal_at = code.index("REFUSED, BEFORE EXCEL WAS STARTED.")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert refusal_at < excel_at
    assert "exit 1" in code[refusal_at:excel_at]


def test_30_the_report_is_written_through_and_separates_prerequisites() -> None:
    body = _function("Write-W6Line")
    assert "Set-Content -LiteralPath $script:W6Path" in body
    code = _own_code()
    assert "'PREREQUISITE'" in code
    assert "$results.Count -gt 0" in code, (
        "a run that recorded no RESULT at all would otherwise pass")








# ===========================================================================
# H. THE FIXTURE AND THE RULE UNDER TEST
# ===========================================================================

def _at(needle: str) -> int:
    code = _own_code()
    assert needle in code, needle
    return code.index(needle)


def _selector() -> dict:
    return json.loads(
        (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8")
    )["selector_semantics"]


def test_31_the_fixture_is_the_accepted_w5_baseline_with_two_selector_rungs() -> None:
    case = _fixture()
    model = case["model"]
    assert len(model["cost_lines"]) + len(model["risks"]) == 5
    assert model["timeline"]["duration"] == 4
    assert case["iterations"] == 1000
    assert case["seed_mode"] == "FIXED"
    assert case["supplied_seed"] == 20260905
    assert case["selected_confidence_level"] == "P80"
    assert case["second_confidence_level"] == "P50"
    code = _own_code()
    # BOTH LABELS COME FROM THE CORPUS. Neither is typed into the runner.
    assert "$firstLabel = [string]$case.selected_confidence_level" in code
    assert "$secondLabel = [string]$case.second_confidence_level" in code
    assert "'P80'" not in code and "'P50'" not in code, (
        "a selector label is typed into the runner rather than read from the corpus")
    assert "the two selector positions are distinct projected ladder rungs" in code
    assert "($firstIndex -ne $secondIndex)" in code
    assert "($firstProbability -ne $secondProbability)" in code


def test_32_the_contract_rule_is_projected_and_asserted_before_it_is_tested() -> None:
    """THE W5 LESSON APPLIED. A semantic nothing projects is a semantic the
    harness has to guess, and W5 guessed wrong. These four booleans are the
    whole of what W6 proves, so they are projected and stated as a
    precondition: if the contract ever changed one, W6 must fail loudly rather
    than quietly test something else."""
    contract = (PCCM_ROOT / "spec" / "sim_contract.yaml").read_text(encoding="utf-8")
    for line in ("distribution_currentness_is_selector_specific: false",
                 "profile_currentness_is_selector_specific: true",
                 "profile_relabelled_on_selector_change: false",
                 "selector_change_requires_new_simulation: false"):
        assert line in contract, line
    semantics = _selector()
    assert semantics["distribution_currentness_is_selector_specific"] is False
    assert semantics["profile_currentness_is_selector_specific"] is True
    assert semantics["profile_relabelled_on_selector_change"] is False
    assert semantics["selector_change_requires_new_simulation"] is False
    assert semantics["selector_input_key"] == "selected_confidence_level"
    # PROJECTED FROM THE CONTRACT, NOT RETYPED IN THE BUILDER.
    generator = (PCCM_ROOT / "builder" / "pccm_builder"
                 / "phase7_acceptance.py").read_text(encoding="utf-8")
    for key in ("distribution_currentness_is_selector_specific",
                "profile_currentness_is_selector_specific",
                "profile_relabelled_on_selector_change",
                "selector_change_requires_new_simulation"):
        assert f'annual["{key}"]' in generator, key
    # AND THE RUNNER ASSERTS IT.
    code = _own_code()
    assert "the contract still says the two annual products differ on selector currentness" in code
    assert "$selectorSemantics.distribution_currentness_is_selector_specific" in code
    assert "$selectorSemantics.profile_currentness_is_selector_specific" in code
    assert "$selectorSemantics.profile_relabelled_on_selector_change" in code
    assert "$selectorSemantics.selector_change_requires_new_simulation" in code


def test_33_the_builder_refuses_a_contract_that_merged_the_two_rules() -> None:
    """If the two products ever shared a currentness rule the scenario would be
    proving nothing, so the emitter refuses it rather than emitting a corpus
    that quietly makes W6 vacuous."""
    from pccm_builder import phase7_acceptance as emitter

    spec = PCCM_ROOT / "spec"
    sim = load_sim_contract(spec / "sim_contract.yaml")
    limits = load_structure_contract(spec / "structure_contract.yaml").limits
    inspection = emitter.build_phase7_inspection(
        sim, load_calc_contract(PCCM_ROOT / "spec" / "calc_contract.yaml"),
        limits.max_generated_year_columns)
    cases = emitter.build_phase7_cases(
        sim=sim, calc=load_calc_contract(spec / "calc_contract.yaml"),
        max_record_rows=limits.max_generated_year_columns,
        min_year=limits.min_year, max_year=limits.max_year)
    emitter.validate_phase7_artifacts(inspection, cases)
    merged = json.loads(json.dumps(inspection))
    merged["selector_semantics"]["profile_currentness_is_selector_specific"] = False
    with pytest.raises(ValueError, match="share"):
        emitter.validate_phase7_artifacts(merged, cases)


# ===========================================================================
# I. THE SELECTOR MOVES, AND NOTHING ELSE RUNS
# ===========================================================================

def test_34_the_selector_is_written_through_the_ordinary_user_input() -> None:
    """NOT A MACHINE CELL. The whole claim is about what the ordinary reporting
    selector does, so it is written through the projected Setup input and
    nothing else is touched."""
    code = _own_code()
    assert code.count("Set-W6NamedText -Workbook $wb") == 2, (
        "the selector is written exactly twice - once for the baseline, once "
        "for the move")
    # BOTH WRITES GO THROUGH THE PROJECTION. One of them doing so is not enough:
    # the move is the one that matters, and it is the second.
    assert code.count(
        "$inspection.inputs.($selectorSemantics.selector_input_key).defined_name") == 2, (
        "a selector write names the input in the runner rather than reading it "
        "from the projection")
    assert not re.search(r"'inp\w+'", code), (
        "a defined-name literal appears in the runner")
    assert "'selected_confidence_level'" not in code
    for banned in ("Set-SimRawCell", "Set-SimField", "$rng.Value2 ="):
        outside = code.replace(_function("Set-W6NamedText"), "")
        assert banned not in outside, banned


def test_35_nothing_runs_between_the_selector_move_and_its_checks() -> None:
    """A rerun hidden between the move and the observation would make Part A
    prove the opposite of what it claims."""
    code = _own_code()
    move_at = _at("PART A - THE SELECTOR MOVES")
    rerun_at = _at("PART B - ANNUAL POST-PROCESSING RERUNS")
    between = code[move_at:rerun_at]
    assert "Invoke-Phase6Simulation" not in between, (
        "the simulation is rerun during the selector-move observation")
    assert "Invoke-W6Annual" not in between, (
        "annual post-processing runs before Part A has finished observing")
    assert "PCCM_Calculate" not in between
    # And the move itself really is inside that window.
    assert "Set-W6NamedText -Workbook $wb `" in between


def test_36_part_a_proves_the_simulation_did_not_move() -> None:
    code = _own_code()
    assert "Add-W6InvariantChecks 'the selector move'" in code
    assert "the simulation is still CURRENT after the selector move" in code
    assert "the active bank did not move" in code
    assert "the selector move changed no published iteration value" in code
    # THE FINGERPRINT SEMANTIC, NAMED AND PROVED - not inferred from a CURRENT.
    move_at = _at("PART A - THE SELECTOR MOVES")
    rerun_at = _at("PART B - ANNUAL POST-PROCESSING RERUNS")
    between = code[move_at:rerun_at]
    for field in ("'request_fingerprint'", "'result_digest'", "'run_id'", "'effective_seed'",
                  "'supplied_seed'", "'iterations_run'", "'consumed_auto_nonce'"):
        assert field in between, field
    assert "s ' +\n                             $field + ' unchanged'" in between


def test_37_part_a_proves_the_distributions_did_not_move() -> None:
    code = _own_code()
    assert "the selector move rewrote nothing in the persisted annual answer" in code
    assert "every persisted distribution rung is value-identical after the move" in code
    assert "the annual stamp is still bound to the same successful simulation" in code
    assert "Compare-W6Surface -Before $annualBefore -After $annualAfterMove -Only '.ladder_'" in code
    # THE SURFACE IS THE WHOLE ANSWER, rung by rung, year by year.
    surface = _function("Get-W6AnnualSurface")
    assert "Get-W6AnnualStamp" in surface and "Get-W6AnnualRecord" in surface
    assert "[int]$P7.annual_records.quantile_count" in surface
    assert "for ($offset = 0; $offset -lt $YearCount; $offset++)" in surface
    compare = _function("Compare-W6Surface")
    assert "Test-SimSameValue" in compare, (
        "the surface comparison does not compare types as well as values")


def test_38_part_a_proves_the_old_profile_is_kept_and_reported_as_other_px() -> None:
    """THE SEMANTIC SPLIT. Distribution CURRENT, profile OTHER Px, old values
    untouched and never relabelled."""
    code = _own_code()
    assert "$otherPxState = [string]$p7.handoff.profile_states[2]" in code
    assert "reports the projected OTHER-Px state" in code
    assert "'OTHER Px'" not in code, (
        "the OTHER-Px state is typed into the runner rather than projected")
    assert "is still CURRENT after the selector move" in code
    assert "still identifies the OLD " in code
    assert "and was not relabelled" in code
    assert "the projected profile vocabulary carries an OTHER-Px state" in code
    projection = json.loads(
        (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8"))
    assert projection["handoff"]["profile_states"] == [
        "NOT PRODUCED", "CURRENT", "OTHER Px", "HISTORICAL"]
    assert projection["handoff"]["distribution_states"][1] == "CURRENT"


# ===========================================================================
# J. THE RERUN
# ===========================================================================

def test_39_part_b_reruns_only_the_annual_endpoint() -> None:
    code = _own_code()
    assert code.count("Invoke-Phase6Simulation") == 1, (
        "the simulation is run more than once; W6 is one simulation and two "
        "selector positions")
    rerun_at = _at("PART B - ANNUAL POST-PROCESSING RERUNS")
    after = code[rerun_at:]
    assert "Invoke-Phase6Simulation" not in after
    assert "PCCM_Calculate" not in after
    assert "$annualSecond = Invoke-W6Annual" in after
    assert "the annual rerun allocated no second run identity" in code
    assert "the annual rerun opened no second publication bank" in code
    assert "the annual rerun consumed no AUTO nonce and moved no pending marker" in code
    assert "the annual rerun changed no published iteration value" in code


def test_40_part_b_proves_the_distributions_are_identical_to_the_baseline() -> None:
    """A REPORTING SELECTOR MAY NOT RECOMPUTE A DISTRIBUTION, and the comparison
    is against the P80 BASELINE - not against the state after the move, which
    would only prove the rerun was idempotent with itself."""
    code = _own_code()
    assert ("Compare-W6Surface -Before $annualBefore -After $annualAfterRerun -Only '.ladder_'"
            in code), "the rerun's ladders are compared against the wrong baseline"
    assert "is value-identical to the " in code and "baseline after the rerun" in code
    assert "the project-year identities are unchanged" in code
    assert "the calendar years are unchanged" in code
    # And the stamp's run-identity half is still the baseline's.
    for field in ("'run_id'", "'effective_seed'", "'request_fingerprint'",
                  "'result_digest'", "'iterations'", "'year_count'"):
        assert field in code, field
    assert "is still the baseline value" in code


def test_41_the_new_profile_must_be_stamped_but_equality_is_not_a_failure() -> None:
    """The stamp MUST move; the values are reported. Failing on equality would
    assert a property of this fixture rather than of the code."""
    code = _own_code()
    assert "the profile is now stamped " in code
    assert "at its projected probability" in code
    assert "Test-SimExactText -Actual $annualAfterRerun['stamp.selected_px_label']" in code
    assert "Test-SimExactDouble -Actual $annualAfterRerun['stamp.selected_px_probability']" in code
    # THE VALUE CHANGE IS REPORTED, NEVER REQUIRED.
    assert "(reported, not required)" in code
    changed = re.search(r"\$profileMoved = Compare-W6Surface[^\n]*\n?[^\n]*", code)
    assert changed, "the profile change is not measured at all"
    following = code[code.index("$profileMoved ="):]
    following = following[:following.index("$handoffAfterRerun")]
    assert "Add-W6Check" not in following, (
        "the profile value change is asserted; a fixture that mathematically "
        "produced equal values would fail for the wrong reason")
    assert "Write-W6Line" in following


def test_42_the_reconciliation_is_one_function_used_at_both_rungs() -> None:
    """A second copy for the second selector would be a second behaviour, and
    the point of W6 is that only the rung changes."""
    code = _own_code()
    assert code.count("Invoke-W6Reconciliation") == 3, (
        "the reconciliation is defined once and invoked once per selector rung")
    body = _function("Invoke-W6Reconciliation")
    assert "-Label $firstLabel -Probability $firstProbability -LadderIndex $firstIndex" in code
    assert "-Label $secondLabel -Probability $secondProbability -LadderIndex $secondIndex" in code
    # THE THREE CHECKS W5 CLOSED ON, unchanged in substance.
    assert "TOTAL equals the contract" in body
    assert "selected-Px profile sums to the reported" in body
    assert "contingency is " in body
    assert "$semantics.contingency_formula" in body
    assert "$semantics.baseline_metric_key" in body
    assert "Get-W6Type7Value" in body
    assert "Get-W6IdentityAllowance -Provenance $Provenance" in body
    # THE ROW KEY FOLLOWS THE RUNG, so P50 reads the P50 rung and not P80's.
    assert "$rowKey = 'quantile_' + [string]($LadderIndex + 1)" in body
    assert "-RowKey $rowKey" in body


def test_43_the_profile_sums_are_read_back_and_never_recomputed() -> None:
    body = _function("Get-W6ProfileSums")
    assert "Get-W6AnnualRecord" in body
    assert "$record[('profile_' + $measure)]" in body
    assert "$value -isnot [double]" in body, "a profile published as text would pass"
    assert "[Math]::Abs([double]$value)" in body, "the conditioning scale is not accumulated"
    code = _own_code()
    for banned in ("AnnualVector", "blend", "lo_hi_f"):
        assert banned not in code, f"{banned} suggests the blend is being rebuilt"
    assert code.count("Get-W6Type7Value") == 2, (
        "the Type-7 helper is defined once and used once, on the iteration "
        "TOTALS column")
    assert not re.search(r"\b\d*\.?\d+e-\d+\b", code), (
        "a numeric tolerance literal appears in the runner")


def test_44_persistence_stays_compact_and_the_measures_stay_projected() -> None:
    code = _own_code()
    assert "the annual region still holds exactly " in code
    assert "the row immediately after the annual result is still blank" in code
    assert "$boundaryRow = $yearCount + 1" in code
    region = _function("Get-W6AnnualRegionIndex")
    assert "[int]$records.max_record_rows" in region
    assert "UsedRange" not in code and "SpecialCells" not in code
    assert "$measures = @($p7.summary_semantics.contingency_measures" in code
    assert "@('nominal', 'pv')" not in code, (
        "a measure list is typed into the runner rather than projected")


def test_45_the_runner_states_what_no_windows_oracle_owns() -> None:
    text = _text()
    assert "THERE IS NO INDEPENDENT ANNUAL WINDOWS ORACLE" in text
    assert "proved independently on Linux" in text
    assert "never recomputed" in text or "never reconstructed" in text or (
        "The blend itself is never recomputed here." in text)
