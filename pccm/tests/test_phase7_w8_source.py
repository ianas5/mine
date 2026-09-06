#!/usr/bin/env python3
"""P7-10: the MINIMAL W8 runner, proved on Linux before a Windows session.

W8 IS THE ONE SCENARIO WHERE NOTHING IS PRODUCED. A refusal is only worth
anything if it costs nothing, so W8 establishes one successful annual answer,
reconciles it, and then reaches two different refusing states through ordinary
user inputs and requires the answer to be exactly where it was afterwards.

TWO STATES, TWO ROUTES, AND THEY ARE GENUINELY DIFFERENT. Part A moves the
SIMULATION REQUEST - the Monte Carlo iteration count, which enters the request
fingerprint and no Phase-5 calculation - so the model stays current and the
published run goes stale. Part B moves the MODEL, through the accepted cost-line
register: one driver's maximum is put below its own minimum, which is the
ordering modCalcCheck refuses, so Phase 5 cannot be prepared at all and no
current request fingerprint exists to compare.

TWO AXES, KEPT APART. A persistent state says what something IS; an attempt
result says what happened when somebody pressed something. REFUSED belongs to
the second and may never appear in the first. The runner projects six
vocabularies, proves the axes disjoint before it uses them, and then proves at
runtime that no accessor and no persisted state cell answers with an attempt
word.

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
RUNNER = WINDOWS / "phase7_w8_refusal.ps1"
W1 = WINDOWS / "phase7_w1_smoke.ps1"
W2 = WINDOWS / "phase7_w2_many_drivers.ps1"
W3 = WINDOWS / "phase7_w3_long_years.ps1"
W4 = WINDOWS / "phase7_w4_base_simulation.ps1"
W5 = WINDOWS / "phase7_w5_annual_success.ps1"
W6 = WINDOWS / "phase7_w6_selector_move.ps1"
W7 = WINDOWS / "phase7_w7_bank_cycle.ps1"
TIMING = WINDOWS / "phase7_timing_scenarios.ps1"
LIFECYCLE = WINDOWS / "com_lifecycle.ps1"
PHASE5 = WINDOWS / "phase5_gate_b_scenarios.ps1"
PHASE6 = WINDOWS / "phase6_gate_b_scenarios.ps1"
BUILD = PCCM_ROOT / "build"
SRC = PCCM_ROOT / "src" / "vba"

DOT_SOURCED = (LIFECYCLE, PHASE5, PHASE6)

COPIED_HELPERS = (
    "Write-RowObject", "Get-NamedValue", "Set-NamedValue", "Get-TableColumnNames",
    "Set-TableCell", "Get-TableBody", "Get-TableRowCount", "Add-BlankTableRow",
    "Remove-TableRow", "Get-IdColumnValues",
)

# SIX OF THE TEN ARE PURELY TRANSITIVE - the accepted fixture calls them and this
# runner never does. FOUR are named as exceptions: W8 reads and writes the
# projected simulation request through the two defined-name primitives, and
# reaches the cost-line register through the two table primitives, which is
# exactly how the accepted fixture writer reaches the same cells.
CALLED_DIRECTLY = ("Get-NamedValue", "Set-NamedValue", "Get-TableBody", "Set-TableCell")

# W1 THROUGH W7 ARE CLOSED. They are accepted Windows evidence and are pinned,
# not maintained, by this round. The digests are LITERAL - hashing a file at
# import time would be a control that can never fail.
W1_SHA256 = "bf41a2193dbf2ea5fbddbe4dad3c6e94649f21262dcf9800ae71eda6e274b59f"
W2_SHA256 = "558bcfa0528b38c28c6e81dd726a18fed9078098f79bcabb792304e6255d8fee"
W3_SHA256 = "97af69bf836a2d798be4937f1e81cfe7218d2cfd702c1b96aa25623e7f8c12be"
W4_SHA256 = "9e1be0629d337b42b02bf790feef752c1fd37fbf0e6c400d4c90af3cd9d86cd0"
W5_SHA256 = "de2e3322c43e93da9647cca0004500e1427e47faf89cdc55c58c7b37e5159ea7"
W6_SHA256 = "feefceb0b7b0592a5799ffd5f95c9ac02c1d7056dfaf04d8d20bf868e7e7f019"
W7_SHA256 = "6abcf0874909c21f6abe7950189e53b72f14a75bb16e0ea3efbaab7ad1de1d99"

# THE PARTS W8 SHARES WITH W7, PINNED SO THE MANDATED REUSE CANNOT DRIFT. The
# instruction was to reuse the W1-W7 lifecycle and the W5/W6 reconciliation
# unchanged; two runners that started identical and quietly diverged would be the
# cost of that, so the shared functions are compared after normalising the name.
SHARED_WITH_W7 = (
    "Write-W8Line", "Add-W8Check", "Invoke-W8Release", "Get-W8SourceRevision",
    "ConvertTo-W8ColumnNumber", "ConvertFrom-W8ColumnNumber", "Get-W8AnnualStamp",
    "Get-W8AnnualRecord", "Get-W8AnnualSurface", "Get-W8BankCapture",
    "Get-W8IterationBlock", "Compare-W8IterationGrid", "Get-W8Handoff",
    "Set-W8NamedText", "Get-W8Type7Value", "Get-W8IdentityAllowance",
    "Get-W8ProfileSums", "Invoke-W8Reconciliation", "Get-W8CandidateBank",
    "Test-W8Authoritative",
)

_CACHE: dict = {}


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    """The runner with its block comment and line comments removed.

    Its header explains what W8 refuses to do and names the things the controls
    forbid, so a scan that read the prose would convict the file of its own
    documentation."""
    return accepted._ps_code(RUNNER)


def _functions() -> dict[str, str]:
    """name -> body, brace-matched over the comment-stripped source.

    String literals are KEPT: the projected keys, the register column name and
    the announcement prefixes are all literals, and they are the subject here."""
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
    assert body is not None, f"{name} is not defined in the W8 runner"
    return body


def _own_code() -> str:
    """The runner's top level plus the functions it wrote itself.

    The ten copied helpers are accepted code from another file; a claim about
    what THIS runner does must not be made about them, and a claim about them is
    made by pinning them instead."""
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
    """THE BEHAVIOURAL CASE. W8 introduces no model of its own: the baseline it
    refuses to damage is the one W4 established and W5 and W6 reconciled."""
    case = [s for s in _cases()["scenarios"] if s["id"] == "W4"]
    assert case, "the acceptance corpus carries no behavioural scenario"
    return case[0]


def _p7() -> dict:
    if "p7" not in _CACHE:
        _CACHE["p7"] = json.loads(
            (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["p7"]


def _sim_inspection() -> dict:
    if "sim" not in _CACHE:
        _CACHE["sim"] = json.loads(
            (BUILD / "phase6_gate_b_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["sim"]


def _gate_b_cases() -> dict:
    if "gb" not in _CACHE:
        _CACHE["gb"] = json.loads(
            (BUILD / "phase6_gate_b_cases.json").read_text(encoding="utf-8"))
    return _CACHE["gb"]


def _manifest() -> dict:
    if "manifest" not in _CACHE:
        _CACHE["manifest"] = json.loads(
            (BUILD / "stage_b_manifest.json").read_text(encoding="utf-8"))
    return _CACHE["manifest"]


def _vocabularies() -> dict[str, list[str]]:
    """Every state and attempt word any of the six projections declares."""
    p7 = _p7()
    gb = _gate_b_cases()["vocabulary"]
    return {
        "sim_states": list(gb["sim_states"]),
        "sim_attempts": list(gb["attempt_results"]),
        "calc_states": list(p7["model_states"]["derived_status"]),
        "calc_attempts": list(p7["model_states"]["attempt_result"]),
        "distribution_states": list(p7["handoff"]["distribution_states"]),
        "profile_states": list(p7["handoff"]["profile_states"]),
    }


# ===========================================================================
# A. SCOPE: THIS IS W8, AND W1 THROUGH W7 ARE CLOSED
# ===========================================================================

def test_01_the_runner_exists_and_is_dedicated() -> None:
    assert RUNNER.exists()
    lines = _text().splitlines()
    assert len(lines) < 2000, (
        f"{len(lines)} lines; W8 rebuilds a baseline, reconciles it and then "
        "proves two refusals - if it has grown past that it is not minimal")
    head = "\n".join(lines[:70])
    assert "W8" in head and "REFUS" in head.upper()


def test_02_it_dot_sources_only_definition_only_files() -> None:
    """THE THREE FILES WHOSE TOP LEVEL DEFINES AND DOES NOTHING. Dot-sourcing
    phase4_functional_test.ps1 or phase7_timing_scenarios.ps1 executes their
    scenarios; that is why the ten helpers are copied instead."""
    sourced = re.findall(r"^\.\s+\(Join-Path \$scriptDir '([^']+)'\)", _code(), re.M)
    assert sourced == ["com_lifecycle.ps1", "phase5_gate_b_scenarios.ps1",
                       "phase6_gate_b_scenarios.ps1"], sourced


def test_03_the_stage_b_bootstrap_is_a_child_script_not_a_dot_source() -> None:
    code = _code()
    assert re.search(r"^\$bootstrap = Join-Path \$scriptDir 'build_stage_b\.ps1'", code, re.M)
    assert re.search(r"^& \$bootstrap -BuildDir \$tempRoot -Force", code, re.M)
    assert ". (Join-Path $scriptDir 'build_stage_b.ps1')" not in code


def test_04_it_stays_inside_its_own_scenario() -> None:
    """W8 IS REFUSAL AND PRESERVATION. It moves no reporting selector, cycles no
    bank, shrinks no duration, runs no sensitivity and produces no second
    successful result of anything."""
    own = _own_code()
    for forbidden in ("PCCM_RunSensitivity", "PCCM_SensitivityState",
                      "second_confidence_level", "shrink_model", "shrink_expected",
                      "second_supplied_seed", "Get-W8SurplusResidue"):
        assert forbidden not in own, f"W8 reaches {forbidden}, which is not its scenario"
    # ONE SUCCESSFUL SIMULATION AND ONE SUCCESSFUL ANNUAL RUN, both in the
    # baseline. A second of either would make W8 a run scenario.
    assert own.count("Invoke-Phase6Simulation") == 1
    # THE ANNUAL ENDPOINT IS INVOKED TWICE - once for the baseline and once
    # inside the shared refusal probe, which runs for each part - and every
    # invocation reads the projected name rather than spelling it.
    assert own.count("-Endpoint ([string]$p7.command_surface.annual_endpoint)") == 1
    assert own.count("-Endpoint ([string]$P7.command_surface.annual_endpoint)") == 1
    assert "PCCM_RunAnnualStochastic" not in own, (
        "the annual endpoint is projected, and typing its name here makes this "
        "runner a second authority for the command surface")


def test_05_it_does_not_repeat_the_w1_surface_matrix() -> None:
    """W1 proved the command surface exists and compiles, once, and is closed.
    W8 asks the workbook one compile question and then gets on with its own."""
    own = _own_code()
    assert own.count("PCCM_CalculationStatus") <= 8
    assert "ProcOfLine" not in own and "CodeModule" not in own


def test_06_the_accepted_runners_are_untouched() -> None:
    """ALL SEVEN ARE CLOSED on accepted evidence. This round does not modify or
    rerun any of them, and an edit must fail here rather than pass."""
    assert all(path.exists() for path in (W1, W2, W3, W4, W5, W6, W7))
    for path, digest, name in ((W1, W1_SHA256, "W1"), (W2, W2_SHA256, "W2"),
                               (W3, W3_SHA256, "W3"), (W4, W4_SHA256, "W4"),
                               (W5, W5_SHA256, "W5"), (W6, W6_SHA256, "W6"),
                               (W7, W7_SHA256, "W7")):
        assert hashlib.sha256(path.read_bytes()).hexdigest() == digest, (
            f"the accepted {name} runner has been modified; it is closed evidence")
    code = _code()
    for closed in ("phase7_w1_smoke", "phase7_w2_many_drivers", "phase7_w3_long_years",
                   "phase7_w4_base_simulation", "phase7_w5_annual_success",
                   "phase7_w6_selector_move", "phase7_w7_bank_cycle"):
        assert closed not in code, f"W8 reaches into the closed {closed} at run time"


def test_06b_the_frozen_large_harness_is_neither_patched_nor_reached() -> None:
    frozen = WINDOWS / "phase7_acceptance_scenarios.ps1"
    assert hashlib.sha256(frozen.read_bytes()).hexdigest() == (
        "9744d9b7c1b4ebbc94ae48db54dd2ffb74af9a554ee43d8da8b1e06efe68c8ed"), (
        "the large acceptance harness is frozen as history and was modified")
    assert "phase7_acceptance_scenarios" not in _code()


def test_06c_the_shape_shared_with_w7_is_identical_to_w7s() -> None:
    ours = _text()
    theirs = W7.read_text(encoding="utf-8")
    for name in SHARED_WITH_W7:
        mine = re.search(rf"^function {name} \{{(.*?)^\}}", ours, re.M | re.S)
        assert mine, f"{name} is not defined in the W8 runner"
        w7name = name.replace("W8", "W7")
        yours = re.search(rf"^function {w7name} \{{(.*?)^\}}", theirs, re.M | re.S)
        assert yours, f"{w7name} is no longer in the accepted W7 runner"
        assert mine.group(1).replace("W8", "W7") == yours.group(1), (
            f"{name} has drifted from the accepted {w7name}; the lifecycle and "
            "the reconciliation were to be reused unchanged")


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
    """TRANSITIVELY, because the dependency that broke W1 was transitive."""
    reached, missing = _closure()
    assert not missing, (
        "custom command(s) called but defined neither in the runner nor in the "
        "files it dot-sources:\n  " +
        "\n  ".join(f"{call}  (reached from {where})" for where, call in sorted(set(missing))))
    assert len(reached) > 40, f"only {len(reached)} custom commands reached"


def test_08_the_copied_helpers_are_verbatim_and_honestly_labelled() -> None:
    timing = TIMING.read_text(encoding="utf-8")
    raw = _text()
    own = _own_code()
    body = "\n".join(l for l in own.splitlines() if not l.strip().startswith("function "))
    for name in COPIED_HELPERS:
        theirs = re.search(rf"^function {name} \{{(.*?)^\}}", timing, re.M | re.S)
        assert theirs, f"{name} is no longer in the accepted timing harness"
        ours = re.search(rf"^function {name} \{{(.*?)^\}}", raw, re.M | re.S)
        assert ours, f"{name} is not defined in the W8 runner"
        assert ours.group(1) == theirs.group(1), (
            f"{name} has drifted from the accepted implementation")
        called = re.search(rf"(?<![\w\-.$]){name}(?![\w\-])", body)
        if name in CALLED_DIRECTLY:
            assert called, (
                f"{name} is listed as called directly but nothing calls it; the "
                "justification in the runner no longer describes it")
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
    assert accepted._join_path_positional_count(_code()) <= 2, (
        "Join-Path with more than one child is PowerShell 6+ only")


def test_10_no_powershell_6_or_7_only_construct_is_used() -> None:
    code = _code()
    for label, pattern in accepted.PS51_ONLY_CONSTRUCTS:
        assert not re.search(pattern, code), f"{label} is not available in Windows PowerShell 5.1"


# ===========================================================================
# D. THE COM LIFECYCLE W1 CLOSED ON
# ===========================================================================

def test_25_every_acquisition_is_counted_and_released_into_one_ledger() -> None:
    code = _code()
    assert code.count("New-ReleaseLedger") == 1, "one ledger, created before Excel exists"
    ledger_at = code.index("New-ReleaseLedger")
    assert ledger_at < code.index("New-Object -ComObject Excel.Application")
    # THE THREE SESSION OBJECTS AND THE RANGES, each counted where it is acquired.
    assert code.count("$comAcquired = $comAcquired + 1") == 3
    assert "$comAcquired = $comAcquired + [int]$iterationsBaseline.Acquired" in code
    assert code.count("$comAcquired = $comAcquired + (Invoke-W8") == 4, (
        "every reader that acquires COM objects must return its count into the "
        "same total, or the acquire/release balance is a guess")
    assert "'every COM object this runner acquired was released'" in code


def test_26_the_lifecycle_verdict_checks_are_all_present() -> None:
    code = _code()
    for claim in ("the owned Excel process exited naturally",
                  "no emergency cleanup was required",
                  "every COM object this runner acquired was released",
                  "every COM release succeeded",
                  "every COM release left 0 outstanding references",
                  "every transient release inside the accepted readers succeeded"):
        assert f"'{claim}'" in code, f"the lifecycle claim {claim!r} is not checked"
    assert "$script:W8Residual" in code and "[int]$rec.Count -ne 0" in code


def test_27_the_shutdown_is_the_accepted_path_and_emergency_is_never_a_pass() -> None:
    code = _code()
    assert "$wb.Close($false)" in code and "$excel.Quit()" in code
    assert "$Error.Clear()" in code
    assert code.count("[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()") >= 3
    assert "Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90" in code
    assert "$emergencyRequired = $true" in code
    assert "'no emergency cleanup was required' (-not $emergencyRequired)" in code


def test_28_no_hidden_reference_pattern_and_nothing_is_saved() -> None:
    code = _code()
    assert ".Save()" not in code and "SaveAs" not in code
    assert "$wb.Close($false)" in code, "the workbook is closed WITHOUT saving"
    for chained in ("$excel.Workbooks.Open", "$wb.Worksheets.Item", "$wb.Names.Item"):
        assert chained not in code, f"{chained} mints an RCW nothing owns"


def test_29_the_tree_is_proved_clean_before_excel_is_started() -> None:
    code = _code()
    assert "Get-W8SourceRevision" in code
    at_check = code.index("Get-W8SourceRevision -RepoRoot $repoRoot")
    at_excel = code.index("New-Object -ComObject Excel.Application")
    assert at_check < at_excel
    body = _function("Get-W8SourceRevision")
    assert "'pccm/src', 'pccm/spec', 'pccm/builder'" in body


def test_30_the_report_is_written_through_and_separates_prerequisites() -> None:
    code = _code()
    assert "Set-Content -LiteralPath $script:W8Path" in _function("Write-W8Line")
    assert "$Kind = 'RESULT'" in _function("Add-W8Check")
    assert "'PREREQUISITE'" in code
    assert "($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0)" in code
    assert "($results.Count -gt 0)" in code, (
        "a run that checked nothing must not be able to pass")


# ===========================================================================
# E. THE STATE WORDS ARE PROJECTED, NOT TYPED
# ===========================================================================

def test_31_no_state_or_attempt_word_is_spelled_in_the_runner() -> None:
    """THE CONTROL THE WHOLE SCENARIO RESTS ON. Every one of these words has a
    projection owner, so a quoted occurrence of one is a second authority for it
    - and a runner that typed the word it expects would pass against a workbook
    that had stopped agreeing with the contract."""
    own = _own_code()
    for family, words in _vocabularies().items():
        for word in words:
            for literal in (f"'{word}'", f'"{word}"'):
                assert literal not in own, (
                    f"the runner spells {literal} ({family}); it is projected, "
                    "and typing it here makes this file a second authority")


def test_32_every_vocabulary_is_read_from_its_own_projection() -> None:
    body = _function("Get-W8Vocabulary")
    for path in ("$GateBCases.vocabulary.sim_states",
                 "$GateBCases.vocabulary.attempt_results",
                 "$P7.model_states.derived_status",
                 "$P7.model_states.attempt_result",
                 "$P7.handoff.distribution_states",
                 "$P7.handoff.profile_states",
                 "$P7.handoff.inconsistent_stamp_state"):
        assert path in body, f"{path} is not read by the projection reader"
    # THE SUPERSEDED WORD IS TAKEN FROM THE END OF BOTH LISTS, not from a
    # position somebody remembered.
    assert "$distribution[$distribution.Count - 1]" in body
    assert "$profile[$profile.Count - 1]" in body
    # AND EACH PRODUCT'S CURRENT WORD COMES FROM ITS OWN LIST. The contract
    # declares two lists because the two products do not share a currentness
    # rule; reading one out of the other would collapse exactly that.
    assert "AnnualCurrent        = [string]$distribution[1]" in body
    assert "AnnualProfileCurrent = [string]$profile[1]" in body
    assert "-ceq ([string]$vocabulary.AnnualProfileCurrent)" in _code()


def test_33_the_ordinals_are_guarded_by_a_shape_check_before_they_are_used() -> None:
    """An ordinal into a contract list is only as good as the list's shape. If a
    contract grew a state, the shape check fails here rather than the runner
    quietly asserting the wrong member."""
    code = _code()
    guard = "'every projected vocabulary has the shape its ordinals are read at'"
    assert guard in code
    assert "@($vocabulary.SimStates).Count -eq 3" in code
    assert "@($vocabulary.CalcStates).Count -eq 4" in code
    assert "@($vocabulary.DistributionStates).Count -eq 3" in code
    assert "@($vocabulary.ProfileStates).Count -eq 4" in code
    assert code.index(guard) < code.index("New-Object -ComObject Excel.Application"), (
        "the shape is checked before Excel is started, because a projection that "
        "changed shape is not a Windows question")


def test_34_a_refused_attempt_is_proved_not_to_be_a_persistent_state() -> None:
    """THE DISTINCTION THE AUTHORISATION ASKS FOR, EXPLICITLY. Six pairs, checked
    as sets: no persistent vocabulary may share a member with an attempt one."""
    code = _code()
    assert "'simulation states vs simulation attempts'" in code
    assert "'model states vs model attempts'" in code
    assert "'annual distribution states vs simulation attempts'" in code
    assert "'annual profile states vs simulation attempts'" in code
    assert "'annual distribution states vs model attempts'" in code
    assert "'annual profile states vs model attempts'" in code
    assert "no persistent state vocabulary shares a member with an attempt " in code
    # And the refusal word itself is proved to live on the attempt axis only.
    assert "$vocabulary.CalcRefused) -List $vocabulary.CalcAttempts" in code
    for state_list in ("$vocabulary.CalcStates", "$vocabulary.SimStates",
                       "$vocabulary.DistributionStates", "$vocabulary.ProfileStates"):
        assert f"-not (Test-W8Member -Value ([string]$vocabulary.CalcRefused) -List {state_list})" \
            in code.replace("\n", " ").replace("`", "") or \
            f"-List {state_list}" in code, f"{state_list} is not checked against the refusal word"
    # THE AXES ARE REALLY DISJOINT IN THE PROJECTION, so the check can pass.
    v = _vocabularies()
    assert not set(v["sim_states"]) & set(v["sim_attempts"])
    assert not set(v["calc_states"]) & set(v["calc_attempts"])
    assert not (set(v["distribution_states"]) | set(v["profile_states"])) & (
        set(v["sim_attempts"]) | set(v["calc_attempts"]))


def test_35_the_builder_projects_the_model_state_axes_and_refuses_a_merged_one() -> None:
    """THE PROJECTION W8 NEEDED AND NOBODY HAD MADE. calc_contract.yaml owns the
    two calculation axes; nothing carried them into a JSON the Windows harness
    reads, so a model status could only be named by typing it. It is projected
    now, and the emitter refuses a projection in which the axes overlap."""
    from pccm_builder import phase7_acceptance as emitter
    states = _p7()["model_states"]
    calc = load_calc_contract(PCCM_ROOT / "spec" / "calc_contract.yaml")
    assert states["derived_status"] == list(calc.derived_status_labels)
    assert states["attempt_result"] == list(calc.attempt_result_labels)
    broken = json.loads(json.dumps(_p7()))
    broken["model_states"]["derived_status"] = list(states["derived_status"]) + [
        states["attempt_result"][2]]
    with pytest.raises(ValueError, match="orthogonal"):
        emitter.validate_phase7_artifacts(broken, _cases())
    empty = json.loads(json.dumps(_p7()))
    empty["model_states"]["attempt_result"] = []
    with pytest.raises(ValueError, match="empty axis"):
        emitter.validate_phase7_artifacts(empty, _cases())


def test_36_the_projection_is_generated_not_hand_edited() -> None:
    from pccm_builder import phase7_acceptance as emitter
    sim = load_sim_contract(PCCM_ROOT / "spec" / "sim_contract.yaml")
    calc = load_calc_contract(PCCM_ROOT / "spec" / "calc_contract.yaml")
    limits = load_structure_contract(PCCM_ROOT / "spec" / "structure_contract.yaml").limits
    rebuilt = emitter.build_phase7_inspection(sim, calc, limits.max_generated_year_columns)
    assert rebuilt == _p7(), (
        "phase7_acceptance_inspection.json is not what its generator produces")


# ===========================================================================
# F. THE BASELINE W8 REFUSES TO DAMAGE
# ===========================================================================

def test_37_the_baseline_is_the_accepted_behavioural_fixture() -> None:
    """W8 INVENTS NO MODEL, NO SEED AND NO SELECTOR. The answer whose survival it
    proves has to be one an accepted round already reconciled."""
    case = _fixture()
    assert case["id"] == "W4"
    assert len(case["model"]["cost_lines"]) + len(case["model"]["risks"]) == 5
    assert int(case["model"]["timeline"]["duration"]) == 4
    assert case["seed_mode"] == "FIXED"
    assert int(case["iterations"]) == int(
        _gate_b_cases()["bounds"]["business_minimum_iterations"])
    assert case["selected_confidence_level"] in _gate_b_cases()["vocabulary"]["quantile_labels"]
    code = _code()
    assert "$cases.scenarios | Where-Object { [string]$_.id -ceq 'W4' }" in code
    assert "'the W8 baseline is the accepted behavioural fixture'" in code
    assert "($driverCount -eq 5) -and ($yearCount -eq 4)" in code


def test_38_the_baseline_only_is_reconciled_and_by_the_accepted_function() -> None:
    """THE ANSWER MUST BE REAL BEFORE ITS SURVIVAL MEANS ANYTHING - and there is
    nothing after a refusal to reconcile, because a refusal produces nothing."""
    code = _code()
    assert code.count("Invoke-W8Reconciliation -Workbook") == 1
    assert "-Stage 'baseline'" in code
    # UNCONDITIONAL, AND CHECKED AS SUCH. Counting the call would let a wrapper
    # like `if ($false) { ... }` keep the count and lose the evidence, so every
    # step W8 must actually take is required to START its own line at the
    # session's indentation - which is what "it runs" looks like in this shell.
    for required in ("Invoke-W8Reconciliation -Workbook",
                     "Invoke-W8PreservationChecks -Workbook",
                     "Invoke-W8RefusalProbe -Excel",
                     "Invoke-W8HandoffChecks -Excel"):
        assert not re.search(rf"^\s*(if|while|foreach|switch)\b[^\n]*{re.escape(required)}",
                             code, re.M), (
            f"{required} is reached through a conditional; a step that may not "
            "run is not evidence that it did")
    assert re.search(r"^    Invoke-W8Reconciliation -Workbook", code, re.M), (
        "the baseline reconciliation is not an unconditional session step")
    body = _function("Invoke-W8Reconciliation")
    assert "Get-W8Type7Value" in body
    assert "$P7.summary_semantics.total_percentile_block" in body or \
        "Get-SimSummaryValue" in body, "the TOTAL comes from the summary block"
    assert "$semantics.baseline_metric_key" in body
    assert "$semantics.contingency_formula" in body
    assert "Get-W8IdentityAllowance" in body
    for invented in ("1e-6", "1e-9", "0.000001", "-Tolerance"):
        assert invented not in _own_code(), f"{invented} is a tolerance this round invented"


def test_39_the_runner_states_what_no_windows_oracle_owns() -> None:
    text = _text()
    assert "THERE IS NO INDEPENDENT ANNUAL WINDOWS ORACLE" in text
    assert "none is claimed" in text


# ===========================================================================
# G. THE TWO ROUTES, AND WHY THEY ARE TWO
# ===========================================================================

def test_40_the_stale_route_is_a_valid_projected_iteration_value() -> None:
    """A STALE SIMULATION IS A PUBLISHED RUN THAT NO LONGER ANSWERS THE QUESTION,
    not a broken one. If the new count were outside the projected bounds the
    request fingerprint could fail to form and part A would be testing part B's
    state under part A's name."""
    code = _code()
    assert "$minimumIterations = [int]$gateBCases.bounds.business_minimum_iterations" in code
    assert "$ceilingIterations = [int]$gateBCases.bounds.max_iterations_representable" in code
    assert "$staleIterations = $minimumIterations + 1" in code
    assert "'the stale request is another VALID iteration count, not a broken one'" in code
    assert "($staleIterations -ge $minimumIterations) -and" in code
    assert "($staleIterations -le $ceilingIterations)" in code
    # WRITTEN THROUGH THE PROJECTED CONTROL, and read back through it.
    assert "$iterationsControl = [string]$simInspection.controls.monte_carlo_iterations.defined_name" \
        in code
    assert "Set-NamedValue -Workbook $wb -DefinedName $iterationsControl" in code
    assert "Get-NamedValue -Workbook $wb -DefinedName $iterationsControl" in code
    assert "inpMonteCarloIterations" not in _own_code(), (
        "the control is named by the projection, never by a literal defined name")


def test_41_the_stale_route_leaves_phase_5_alone() -> None:
    code = _code()
    assert "'part A: the model is still the projected current state'" in code
    assert "-ceq ([string]$vocabulary.CalcCurrent)" in code
    assert "'part A: the simulation reports the projected stale state'" in code
    assert "-ceq ([string]$vocabulary.SimStale)" in code


def test_42_the_invalid_route_is_a_normal_register_edit_that_violates_the_ordering() -> None:
    """THE PRODUCTION RULE IT RELIES ON, quoted from the checker that owns it.
    Setting a maximum below the driver's own minimum breaks Min <= Max for a
    two-point family and Min <= Most Likely <= Max for a three-point one, so the
    numerical checker refuses and Phase 5 cannot be prepared at all."""
    checker = (SRC / "modCalcCheck.bas").read_text(encoding="utf-8")
    assert "requires Min <= Max" in checker
    assert "requires Min <= Most Likely <= Max" in checker
    case = _fixture()
    victim = case["model"]["cost_lines"][0]
    assert float(victim["min_value"]) - 1.0 < float(victim["min_value"])
    assert float(victim["min_value"]) - 1.0 < float(victim["most_likely"])
    code = _code()
    assert "$invalidMaximum = $victimMinimum - 1.0" in code
    assert "'the invalidating edit puts a cost line maximum below its own minimum'" in code
    assert "($invalidMaximum -lt $victimMinimum)" in code
    # THE CELL IS REACHED THE WAY THE ACCEPTED FIXTURE REACHES IT: the manifest's
    # own register, its own column list, and the row found by permanent id.
    register = [r for r in _manifest()["registers"] if r["key"] == "cost_lines"]
    assert register and "unit_cost_max" in register[0]["columns"]
    assert "Get-W8Register -Manifest $manifest -Key 'cost_lines'" in code
    assert "-ColumnKey 'unit_cost_max'" in code
    assert "[array]::IndexOf(@($Register.columns), $ColumnKey)" in _function(
        "Get-W8RegisterColumnIndex")
    assert "$PermanentId" in _function("Get-W8RegisterRowIndex")


def test_43_the_invalid_route_recalculates_through_the_accepted_workflow() -> None:
    code = _code()
    assert "Invoke-W8Endpoint -Excel $excel -Endpoint 'PCCM_Calculate'" in code
    assert "'part B: PCCM_Calculate refuses the edited model'" in code
    assert "'part B: the model no longer reports the projected current state'" in code
    assert "-cne ([string]$vocabulary.CalcCurrent)" in code
    assert "-ceq ([string]$vocabulary.CalcInvalid)" in code
    assert "'part B: the simulation reports the projected invalid state'" in code
    assert "-ceq ([string]$vocabulary.SimInvalid)" in code


def test_44_nothing_hidden_is_written_by_the_runner() -> None:
    """BOTH STATES ARE REACHED THE WAY A USER REACHES THEM. Not one cell of
    `_Calc` or `_SimData` is written, no stamp is edited and no INVALID is
    faked."""
    own = _own_code()
    for forbidden in ("Set-SimRawCell", "Set-SimField", "Set-SimPending",
                      "Set-CalcScalar", "Set-Phase5TypedCell"):
        assert forbidden not in own, f"the runner writes hidden state through {forbidden}"
    # The only writes are: the projected simulation request, the projected
    # selector, the accepted fixture, and one register cell.
    writes = re.findall(r"(?<![\w\-.])(Set-[\w]+)(?![\w\-])", own)
    assert set(writes) <= {"Set-NamedValue", "Set-W8NamedText", "Set-TableCell",
                           "Set-Phase5Fixture", "Set-StrictMode", "Set-Content"}, sorted(set(writes))
    assert own.count("Set-TableCell -Workbook $wb") == 1, (
        "exactly one register cell is edited, and it is the invalidating bound")


# ===========================================================================
# H. WHAT A REFUSAL MUST COST: NOTHING
# ===========================================================================

def test_45_both_refusals_are_required_and_must_name_their_state() -> None:
    code = _code()
    assert code.count("Invoke-W8RefusalProbe -Excel $excel") == 2
    assert "-Stage 'part A' -ExpectedSimState ([string]$vocabulary.SimStale)" in code
    assert "-Stage 'part B' -ExpectedSimState ([string]$vocabulary.SimInvalid)" in code
    probe = _function("Invoke-W8RefusalProbe")
    assert "' REFUSED') ($announcement -like 'FAIL|*')" in probe, (
        "a refusal is production's announced FAIL, not the runner's opinion")
    assert "$announcement -clike ('*' + $ExpectedSimState + '*')" in probe, (
        "any failure at all would otherwise count as the right refusal")
    assert "-like 'OK|*'" not in probe


def test_46_the_frozen_partition_comes_from_the_projected_groups() -> None:
    """NOT A LIST TYPED HERE. Every run-identity row carries a group in the
    projection, and the runner excludes exactly the derived group - because the
    status evaluation the annual precondition performs owns those two rows and
    nothing else."""
    body = _function("Get-W8RunInvariants")
    assert "Get-W8RowsInGroup -Inspection $Inspection -Group 'derived'" in body
    assert "$state['pending_auto_nonce']" in body
    assert "$Inspection.publication.bank_labels" in body
    groups = _sim_inspection()["sim_data"]["run_identity"]["groups"]
    assert sorted({k for k, v in groups.items() if v == "derived"}) == [
        "simulation_status", "status_evaluated_at"]
    assert any(v == "attempt" for v in groups.values())
    assert any(v == "snapshot" for v in groups.values())
    reader = _function("Get-W8RowsInGroup")
    assert "$identity.groups.$key) -ceq $Group" in reader
    code = _code()
    assert "'the run-identity projection separates derived rows from attempt rows " in code


def test_47_the_attempt_rows_are_frozen_and_the_claim_is_checked_not_assumed() -> None:
    """THE ANNUAL ENDPOINT SAYS IT WRITES NO ATTEMPT ROW. W8 treats that as a
    claim: the rows stay in the frozen set AND are compared separately, so a
    refusal that wrote one fails rather than being excused."""
    runner = (SRC / "modSimAnnualRun.bas").read_text(encoding="utf-8")
    assert "writes no attempt row" in runner
    probe = _function("Invoke-W8RefusalProbe")
    assert "': the refusal wrote no SIMULATION attempt row'" in probe
    assert "StartsWith('attempt.')" in probe
    metadata = _function("Get-W8StatusMetadata")
    assert "@('attempt', 'derived')" in metadata


def test_48_the_status_row_is_found_by_what_it_holds() -> None:
    """AND ONLY THE TIMESTAMP BESIDE IT MAY MOVE. The row is identified by
    carrying a member of the projected persistent vocabulary - which is also
    what makes 'no state cell answers with an attempt word' a testable claim."""
    probe = _function("Invoke-W8RefusalProbe")
    assert "': exactly one derived row carries a simulation state word'" in probe
    assert "Get-W8RowsHolding -Block $metadataAfter -List $Vocabulary.SimStates" in probe
    assert "-List $Vocabulary.SimAttempts" in probe
    assert "$statusRows.Count -eq 1" in probe
    holder = _function("Get-W8RowsHolding")
    assert "Test-W8Member -Value $Block[$name] -List $List" in holder
    member = _function("Test-W8Member")
    assert "$Value -isnot [string]" in member, (
        "a state word published as a number is not a state word")
    assert "-ceq" in member, "membership is binary, never case-insensitive"


def test_49_physical_persistence_and_semantic_authority_are_proved_separately() -> None:
    """TWO QUESTIONS, TWO FUNCTIONS. The stored payload is compared directly -
    the whole simulation block, the whole annual surface, every published
    iteration value - and the accessors are asked separately what it now MEANS.
    Proving one from the other is exactly what W8 must not do."""
    payload = _function("Invoke-W8PreservationChecks")
    meaning = _function("Invoke-W8HandoffChecks")
    assert "Get-W8BankCapture" in payload and "Compare-W8IterationGrid" in payload
    assert "Get-W8Handoff" not in payload and "PCCM_Annual" not in payload
    assert "Get-W8Handoff -Excel $Excel" in meaning
    assert "Get-W8BankCapture" not in meaning and "Compare-W8Surface" not in meaning
    capture = _function("Get-W8BankCapture")
    assert "Get-SimBankBlock" in capture and "Get-W8AnnualSurface" in capture
    surface = _function("Get-W8AnnualSurface")
    assert "Get-W8AnnualStamp" in surface and "Get-W8AnnualRecord" in surface
    assert "ladder_" in surface, "every rung of both ladders is part of the payload"
    code = _code()
    # FOUR TIMES: after each state change and after each refusal.
    assert code.count("Invoke-W8PreservationChecks -Workbook") == 3
    assert code.count("Invoke-W8HandoffChecks -Excel $excel") == 4


def test_50_the_accessors_are_required_to_report_the_projected_superseded_state() -> None:
    """AND THE ANSWER IS NOT RETRACTED. A published profile keeps the Px it was
    computed at and the year count it covers; a state that stopped being current
    blanks neither and relabels nothing."""
    code = _code()
    assert code.count(
        "-ExpectedDistribution ([string]$vocabulary.DistributionSuperseded)") == 4
    assert code.count("-ExpectedProfile ([string]$vocabulary.ProfileSuperseded)") == 4
    meaning = _function("Invoke-W8HandoffChecks")
    assert "Test-W8Member -Value $distribution -List $Vocabulary.DistributionStates" in meaning
    assert "-not (Test-W8Member -Value $distribution -List $Vocabulary.SimAttempts)" in meaning
    assert "the stamped Px and year count are still the baseline" in meaning
    # THE PRODUCTION RULE THE EXPECTATION PROJECTS FROM.
    store = (SRC / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    assert "SIM_ANNUAL_STATE_HISTORICAL" in store
    assert "ProfileStateOf = distribution" in store, (
        "the profile inherits the distribution state when it is not current")
    assert _vocabularies()["distribution_states"][-1] == \
        _vocabularies()["profile_states"][-1], (
        "the two handoff lists no longer end with the same superseded word")


def test_51_no_publication_may_appear_anywhere_else() -> None:
    payload = _function("Invoke-W8PreservationChecks")
    assert "no annual publication marker appeared in bank " in payload
    assert "$P7.annual_records.stamp.published_marker" in payload
    assert "$Baseline.OtherBank" in payload
    other = _function("Get-W8OtherBank")
    assert "$P7.publication_semantics.bank_labels" in other, (
        "the other bank is the projection's other label, not 'the one that is not A'")


def test_52_nothing_is_allocated_across_the_whole_session() -> None:
    code = _code()
    assert "'no refusal consumed a run id, an AUTO nonce or the pending marker'" in code
    assert "$stateEnd['shared']['next_auto_nonce']" in code
    assert "$stateEnd['pending_auto_nonce']" in code
    assert "$stateEnd['shared']['last_run_id']" in code
    assert "$stateEnd['shared']['active_bank']" in code


def test_53_the_invariants_are_measured_from_the_baseline_not_only_from_before() -> None:
    """HISTORICAL PRESERVATION IS MEASURED AGAINST THE SUCCESSFUL RUN. Comparing
    each step only with the step before it would let a slow drift through, one
    unchanged comparison at a time."""
    probe = _function("Invoke-W8RefusalProbe")
    assert "$Baseline.Invariants" in probe
    assert "measured from the baseline" in probe
    payload = _function("Invoke-W8PreservationChecks")
    assert "$Baseline.Capture" in payload and "$Baseline.Iterations" in payload
    assert "$Baseline.OtherCapture" in payload
    code = _code()
    assert "Invariants   = (Get-W8RunInvariants -Workbook $wb -Inspection $simInspection)" in code
