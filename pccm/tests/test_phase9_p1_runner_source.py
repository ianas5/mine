#!/usr/bin/env python3
"""The Phase-9 Windows runner, checked as far as Linux can check a runner.

WHAT THIS CAN AND CANNOT SETTLE. It cannot say the runner passes: Excel has not
run and this file never pretends otherwise. What it settles is everything that
cost the P8-Z turns before any assertion was reached - a variable read before it
was set, a helper copied without the script state it reads, a PowerShell-7
construct in a file that must run under 5.1, a typed address that a projection
already owns, a startup that could not be walked at all.

AND ONE THING MORE, BECAUSE IT IS THE POINT OF THE FILE. A Phase-9 runner that
re-tested Phase 8 would be re-opening a closed phase by accident. The controls
below refuse the chart, annual and sensitivity vocabulary outright.
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pytest

PCCM_ROOT = Path(__file__).resolve().parents[1]
WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
RUNNER = WINDOWS / "phase9_p1_model_check.ps1"
AUDIT = PCCM_ROOT / "tests" / "powershell_uninitialised_audit.ps1"
PWSH = "/opt/pwsh/pwsh"

sys.path.insert(0, str(PCCM_ROOT / "builder"))


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    """The runner with its commentary removed.

    A BAN ON A TOKEN MUST NOT CONVICT THE SENTENCE EXPLAINING WHY THE TOKEN IS
    BANNED. Every control below that forbids something reads this, not the file.
    """
    lines = []
    for line in _text().splitlines():
        stripped = line.strip()
        if stripped.startswith("#") or stripped.startswith("<#") or stripped == "#>":
            continue
        lines.append(line)
    body = "\n".join(lines)
    # The comment-based help block at the top is one long block comment.
    return re.sub(r"<#.*?#>", "", body, flags=re.S)


# ===========================================================================
# A. IT IS A FILE, IT PARSES, AND IT SAYS WHAT IT IS
# ===========================================================================
def test_01_the_runner_exists_and_declares_that_it_has_not_run() -> None:
    assert RUNNER.is_file()
    text = _text()
    assert "THIS FILE HAS NOT RUN" in text, (
        "the runner does not say that no Windows run has produced it")
    assert "Set-StrictMode -Version 2.0" in text
    assert "$ErrorActionPreference = 'Stop'" in text


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_02_the_runner_parses() -> None:
    script = (
        '$errors = $null; $tokens = $null; '
        f'$null = [System.Management.Automation.Language.Parser]::ParseFile("{RUNNER}", '
        '[ref]$tokens, [ref]$errors); '
        'if ($errors.Count -gt 0) { $errors | ForEach-Object { $_.Message }; exit 1 } '
        'else { "PARSE CLEAN" }')
    done = subprocess.run([PWSH, "-NoProfile", "-Command", script],
                          capture_output=True, text=True, timeout=300)
    assert done.returncode == 0, done.stdout + done.stderr
    assert "PARSE CLEAN" in done.stdout


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_03_no_variable_can_be_read_before_it_is_assigned() -> None:
    """THE CLASS THAT COST P8-Z RUN 2, closed by the accepted AST audit."""
    done = subprocess.run([PWSH, "-NoProfile", "-File", str(AUDIT), "-Path", str(RUNNER)],
                          capture_output=True, text=True, timeout=300)
    assert done.returncode == 0, done.stdout.strip()
    assert done.stdout.strip() == "CLEAN", done.stdout


# ===========================================================================
# B. WINDOWS POWERSHELL 5.1, NOT 7
# ===========================================================================
_PS7_ONLY = (
    "??", "?.", "&&", "||", "-Parallel", "-LeafBase", "-AsHashtable",
    "$IsWindows", "$PSStyle", "ForEach-Object -Parallel",
)


def test_04_nothing_here_needs_powershell_7() -> None:
    """The target is Windows PowerShell 5.1. A 7-only construct fails at parse
    time on the machine that matters, and there is no machine here to find out."""
    code = _code()
    offenders = [token for token in _PS7_ONLY if token in code]
    assert not offenders, f"the runner uses PowerShell 7 constructs: {offenders}"
    # Join-Path takes at most two positional arguments on 5.1.
    for call in re.findall(r"Join-Path [^\n]*", code):
        assert call.count("$") <= 2 or "-Path" in call or "-ChildPath" in call, call


def test_05_the_forbidden_constructs_control_is_not_vacuous() -> None:
    poisoned = _code() + "\n$x = $a ?? $b\n"
    offenders = [token for token in _PS7_ONLY if token in poisoned]
    assert offenders == ["??"], offenders


# ===========================================================================
# C. NOT ONE ADDRESS, NAME OR THRESHOLD IS TYPED
# ===========================================================================
def _string_literals() -> list[str]:
    """Every quoted literal in the runner, scanned LINE BY LINE.

    THE SCAN USED TO PAIR QUOTES ACROSS THE WHOLE FILE and that made it partly
    vacuous: one unbalanced apostrophe anywhere - in a message, in a word like
    "runner's" - desynchronised the alternation and swallowed everything after
    it until the next quote. A literal address in the swallowed region would
    have passed unseen, which is exactly what this control exists to stop.
    Pairing per line bounds any desync to the line that caused it, and there are
    no here-strings in this file for a line-wise scan to break.
    """
    found: list[str] = []
    for line in _code().splitlines():
        found += [single or double for single, double
                  in re.findall(r"'([^']*)'|\"([^\"]*)\"", line)]
    return found


def test_06_no_worksheet_address_is_spelled_in_the_runner() -> None:
    """P7-4 IS WHAT A TYPED ADDRESS COSTS: the persisted block moved, the
    formula did not, and a sheet reported "not produced" after a run that had
    just succeeded. Every address here comes from a projection."""
    # THE ACCEPTANCE SCENARIO IDS ARE NOT ADDRESSES, and they are read from the
    # corpus rather than listed here - a typed exclusion would be a second name
    # for something the corpus already owns.
    import json

    corpus = PCCM_ROOT / "build" / "phase7_acceptance_cases.json"
    allowed = set()
    if corpus.is_file():
        allowed = {str(s["id"]) for s in json.loads(corpus.read_text(encoding="utf-8"))["scenarios"]}
    typed = [value for value in _string_literals()
             if re.fullmatch(r"\$?[A-Z]{1,3}\$?\d{1,5}", value) and value not in allowed]
    assert not typed, f"the runner spells worksheet addresses: {sorted(set(typed))}"


def test_06b_the_address_scan_is_not_vacuous() -> None:
    """SO THE SCAN IS PROVED TO SEE. An address literal injected anywhere - even
    after an apostrophe that used to blind it - has to be found."""
    injected = _code() + "\n$x = 'D8'\n$y = \"the runner's own note\" ; $z = '$AB$1234'\n"
    found = []
    for line in injected.splitlines():
        found += [single or double for single, double
                  in re.findall(r"'([^']*)'|\"([^\"]*)\"", line)]
    hits = {value for value in found if re.fullmatch(r"\$?[A-Z]{1,3}\$?\d{1,5}", value)}
    assert {"D8", "$AB$1234"} <= hits, sorted(hits)


def test_07_no_defined_name_is_spelled_in_the_runner() -> None:
    code = _code()
    for name in ("inpMonteCarloIterations", "inpRandomSeed", "inpSelectedConfidenceLevel"):
        assert name not in code, f"the runner spells the defined name {name}"
    assert "controls.monte_carlo_iterations.defined_name" in code
    assert "controls.random_seed.defined_name" in code


def test_08_the_threshold_is_the_projections_and_never_the_runners() -> None:
    from pccm_builder import load_contract
    from pccm_builder.phase9_model_check import ADVISORY_INPUT_KEY

    threshold = load_contract(PCCM_ROOT / "spec" / "input_contract.yaml").recommendation_for(
        ADVISORY_INPUT_KEY)
    assert str(threshold) not in _code(), "the runner types the advisory threshold"
    assert "$projection.advisory.threshold" in _code()


def test_09_every_vocabulary_word_comes_from_the_projection() -> None:
    """A runner that typed PASS, WARNING or ERROR would be a second declaration
    of a vocabulary the projection already carries - except where it is naming
    the state words the OWNERS publish, which are not Model Check's to declare."""
    code = _code()
    # SCOPED TO A COMPARISON, NOT TO THE WORD. `Add-P9Check` prints its own
    # PASS/FAIL verdict for the runner's checks, which is this file's vocabulary
    # and not Model Check's. What must never happen is a Model Check value being
    # compared against a word typed here.
    for comparison in re.findall(r"-c(?:eq|ne)\s+'([A-Z ]+)'", code):
        assert comparison in ("CURRENT", "INVALID", "STALE", "NOT CALCULATED"), (
            f"the runner compares a Model Check value against the typed word "
            f"{comparison!r}, which the projection owns")
    assert "vocabulary.overall_states" in code
    assert "vocabulary.informational_severity" in code
    assert "vocabulary.severity_order" in code
    # THE OWNERS' OWN WORDS ARE DIFFERENT AND ARE ALLOWED. CURRENT and INVALID
    # belong to modCalcReport and modSimReport, and the runner is asserting what
    # THEY said, not declaring a Model Check vocabulary.
    assert "'CURRENT'" in code and "'INVALID'" in code


# ===========================================================================
# D. IT IS A PHASE-9 RUNNER AND NOTHING ELSE
# ===========================================================================
_PHASE_8_VOCABULARY = ("tornado", "histogram", "s-curve", "scurve", "chartobject",
                       "seriescollection", "reconciliation", "spearman", "rho",
                       "annual profile", "dashboard")


def test_10_it_re_runs_no_phase_8_acceptance() -> None:
    lowered = _code().lower()
    offenders = [word for word in _PHASE_8_VOCABULARY if word in lowered]
    # ChartObjects appears once and only to assert the Model Check sheet draws
    # none - which is a Phase-9 claim about a Phase-9 sheet.
    allowed = {"chartobject"}
    offenders = [word for word in offenders if word not in allowed]
    assert not offenders, f"the Phase-9 runner reaches into Phase-8 territory: {offenders}"
    # ChartObjects APPEARS ONLY TO ASSERT THERE ARE NONE. A runner that read a
    # series, a point or a category would be inspecting a chart.
    for reader in (".SeriesCollection", ".Points(", ".Chart.", ".ChartTitle",
                   ".XValues", ".Values"):
        assert reader not in _code(), f"the runner inspects a chart through {reader}"
    assert "the sheet draws no chart" in _code()
    assert "THIS RUNNER ANSWERS PHASE-9 QUESTIONS ONLY" in _text()


def test_11_it_writes_no_model_check_cell_and_saves_nothing() -> None:
    """THE SURFACE UNDER TEST IS NOT EDITED BY THE TEST. Everything the runner
    changes is an INPUT - a register cell, an iteration count, a seed - and the
    Model Check sheet is only ever read."""
    code = _code()
    assert "$wb.Save()" not in code and ".SaveAs(" not in code
    assert "$wb.Close($false)" in code, "the runner does not close without saving"
    # The only writes are through the two input helpers and production endpoints.
    writes = re.findall(r"\.Value2\s*=", code)
    assert len(writes) <= 3, f"the runner writes cells directly {len(writes)} times"
    for helper in ("Set-NamedValue", "Set-P9TableCell"):
        assert helper in code
    # AND NOTHING IS WRITTEN TO THE SHEET IT READS.
    assert "$modelCheckSheet -Address" not in code.replace("Get-P9Block", "READ")


def test_12_the_states_are_reached_through_production() -> None:
    code = _code()
    for endpoint in ("PCCM_Calculate", "PCCM_RunSimulation"):
        assert endpoint in code, f"the runner never invokes {endpoint}"
    # NO STATE WORD IS EVER WRITTEN. A runner that set a status cell would be
    # asserting against its own fixture.
    assert "WriteStatusBlock" not in code
    assert "PCCM_SimulationStatus" not in code, (
        "the runner calls the persisting simulation entry point")


# ===========================================================================
# E. THE STARTUP REFUSALS
# ===========================================================================
def test_13_it_refuses_a_modified_tree_before_starting_excel() -> None:
    code = _code()
    refusal = code.index("REFUSED, BEFORE EXCEL WAS STARTED.")
    creation = code.index("New-Object -ComObject Excel.Application")
    assert refusal < creation, "the clean-tree refusal is after Excel is created"
    assert "git -C $RepoRoot status --porcelain" in code
    for area in ("pccm/src", "pccm/spec", "pccm/builder"):
        assert area in code, f"the refusal does not cover {area}"


def test_14_it_refuses_a_missing_artefact_by_name() -> None:
    code = _code()
    assert "phase9_model_check_inspection.json" in code
    assert "does not exist. Build Stage A first." in code
    missing = code.index("does not exist. Build Stage A first.")
    creation = code.index("New-Object -ComObject Excel.Application")
    assert missing < creation


def test_15_the_repository_root_is_derived_and_never_machine_specific() -> None:
    code = _code()
    assert "$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path" in code
    assert "$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)" in code
    assert "$repoRoot = Split-Path -Parent $pccmRoot" in code
    assert "C:\\" not in code, "the runner carries a machine-specific path"
    assert "Get-Location" not in code, "the runner depends on the working directory"


# ===========================================================================
# F. THE COM LIFECYCLE
# ===========================================================================
def test_16_every_cleanup_variable_is_assigned_before_the_try() -> None:
    """STRICTMODE BUGS SURFACE ON THE FAILURE PATH, which is the path least
    likely to have been exercised. Every handle and counter the finally block
    reads has to exist before the try is entered."""
    code = _code()
    head = code[:code.index("try {\n    $excel = New-Object")]
    for variable in ("$rel = New-ReleaseLedger", "$excel = $null", "$workbooks = $null",
                     "$wb = $null", "$excelIdentity = $null", "$naturalExit = $false",
                     "$emergencyRequired = $false", "$fatal = ''", "$comAcquired = 0"):
        assert variable in head, f"{variable} is not assigned before the try"
    assert "$preExisting = @(Get-PreExistingExcelPids)" in head


def test_17_every_com_acquisition_is_counted_and_released() -> None:
    code = _code()
    assert code.count("$comAcquired = $comAcquired + 1") == 3, (
        "the acquisition count does not match the three top-level handles")
    for label in ("'Workbook'", "'Workbooks'", "'Application'"):
        assert f"-Label {label}" in code, f"{label} is never released"
    assert "Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90" in code
    assert "Invoke-EmergencyExcelCleanup" in code
    assert code.count("[System.GC]::Collect()") >= 4


def test_18_the_helpers_it_copied_brought_their_script_state_with_them() -> None:
    """THE P8-Z RUN-2 LESSON AS A CONTROL. Every $script: variable the file reads
    is initialised at file scope before anything can read it - which the AST
    audit proves in general and this states in particular, so a reader can see
    which four they are."""
    code = _code()
    head = code[:code.index("function Write-P9Line")]
    for variable in ("$script:P9Lines", "$script:P9Path", "$script:P9Checks",
                     "$script:P9Residual", "$script:P9ErrorCodes"):
        assert f"{variable} = " in head, f"{variable} is not initialised at file scope"


def test_19_it_does_not_dot_source_a_scenario_file() -> None:
    """A SCENARIO FILE IS NOT A LIBRARY. Dot-sourcing the P8-Z runner would run a
    Phase-8 acceptance inside a Phase-9 runner, which is why the six workbook
    helpers are defined here instead."""
    code = _code()
    sourced = re.findall(r"^\. \(Join-Path \$scriptDir '([^']+)'\)", code, re.M)
    assert sorted(sourced) == ["com_lifecycle.ps1", "phase5_gate_b_scenarios.ps1",
                               "phase6_gate_b_scenarios.ps1"], sourced
    for helper in ("function Get-NamedValue", "function Set-NamedValue",
                   "function Set-P9TableCell", "function Add-P9BlankTableRow",
                   "function Remove-P9TableRow", "function Get-P9Register"):
        assert helper in code, f"{helper} is used but not defined here"


# ===========================================================================
# G. THE QUESTIONS ONLY EXCEL CAN ANSWER ARE ACTUALLY ASKED
# ===========================================================================
def test_20_it_asks_whether_the_accessors_answer_from_a_cell() -> None:
    code = _code()
    assert "Test-P9ReadingsAnswered" in code
    assert "without an error" in code
    assert "$script:P9ErrorCodes" in code, "the runner cannot recognise an Excel error"
    assert "-2146826273" in code, "the runner does not know the #VALUE! code"


def test_21_it_asks_whether_the_anchor_makes_the_report_live() -> None:
    """THE SINGLE MOST IMPORTANT WINDOWS QUESTION. PCCM_StructuralReport is not
    volatile; the sheet anchors its cell to the volatile heartbeat instead. If
    that does not work, a fault raised after the workbook opened never appears."""
    code = _code()
    assert "the anchored structural report re-evaluated on an ordinary recalculation" in code
    assert "Add-P9BlankTableRow" in code, "no fault is raised after the workbook opened"
    assert "$Excel.Calculate()" in code, "the runner never asks for a calculation cycle"
    assert "CalculateFull" not in code, (
        "a full rebuild would re-run everything and prove nothing about the anchor")


def test_22_it_asks_whether_the_surface_persists_anything() -> None:
    code = _code()
    assert "recalculating the workbook rewrote no persisted status cell" in code
    assert "calc_state" in code and "calculation_status" in code
    assert "status_evaluated_at" in code


def test_23_it_asks_the_scenario_matrix_and_names_each_scenario() -> None:
    text = _text()
    for scenario in ("SCENARIO A", "SCENARIO B", "SCENARIO C", "SCENARIO D AND G",
                     "SCENARIO E AND M", "SCENARIO F"):
        assert scenario in text, f"the runner never reaches {scenario}"
    code = _code()
    assert "Test-P9Reconciles" in code and "Test-P9WindowShape" in code
    # THE UNUSED SLOTS ARE CHECKED EVERY TIME, not once.
    assert code.count("Test-P9WindowShape -Surface") >= 5


def test_24_the_overflow_disclosure_is_checked_against_the_template() -> None:
    code = _code()
    assert "register.disclosure_template" in code
    assert "the overflow is disclosed in the contracted wording" in code
    assert "nothing is disclosed while the population fits" in code
    assert "Showing the first" not in code, (
        "the runner types the disclosure wording the projection owns")


# ===========================================================================
# H. THE STARTUP, EXECUTED
# ===========================================================================
def _dry_run(source: str) -> tuple[int, str]:
    """Execute the runner's startup, under StrictMode, up to the deliberate
    boundary immediately before Excel is created.

    IT RUNS FROM bootstrap/windows, because $scriptDir is what everything else
    derives from. ONE STEP IS STUBBED AND ONLY ONE: build_stage_b.ps1 drives
    Excel to inject the VBA.
    """
    excel = source.index("    $excel = New-Object -ComObject Excel.Application")
    head = source[:source.rindex("try {", 0, excel)]
    assert "New-Object -ComObject" not in head, "Excel is created inside the startup path"
    head = head.replace(
        "& $bootstrap -BuildDir $tempRoot -Force\n$bootstrapExit = $LASTEXITCODE",
        "if (-not (Test-Path -LiteralPath $bootstrap)) "
        "{ throw 'build_stage_b.ps1 is not beside the runner' }\n"
        "Set-Content -LiteralPath $stageBPath -Value 'dry run' -Encoding UTF8\n"
        "$bootstrapExit = 0")
    # THE CLEAN-TREE REFUSAL IS NEUTRALISED HERE, because this control is about
    # whether the startup can be WALKED, not about whether this checkout happens
    # to be committed. The refusal is real and is asserted in test_13.
    head = head.replace("if ($revision.Dirty.Count -gt 0) {",
                        "if ($false -and $revision.Dirty.Count -gt 0) {")
    head += "\nWrite-Output 'REACHED THE EXCEL BOUNDARY'\nexit 0\n"
    scratch = WINDOWS / "__p9_dryrun_tmp.ps1"
    scratch.write_text(head, encoding="utf-8")
    try:
        done = subprocess.run([PWSH, "-NoProfile", "-File", str(scratch)],
                              cwd="/tmp", capture_output=True, text=True, timeout=300)
    finally:
        scratch.unlink(missing_ok=True)
    return done.returncode, (done.stdout + done.stderr)


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_25_the_startup_reaches_the_excel_boundary_under_strict_mode() -> None:
    """EVERY STARTUP STEP, EXECUTED: the roots derived, the revision captured,
    the refusals evaluated, every artefact resolved and read, the working copy
    made, the report path set, and every script-scope variable initialised."""
    if not (PCCM_ROOT / "build" / "phase9_model_check_inspection.json").is_file():
        pytest.skip("Stage A has not been built into pccm/build")
    code, out = _dry_run(_text())
    assert code == 0, f"the startup did not reach the boundary:\n{out[-3000:]}"
    assert "REACHED THE EXCEL BOUNDARY" in out, out[-2000:]
    for evidence in ("THE MODEL CHECK SURFACE IN REAL EXCEL", "source revision   : ",
                     "row window        : ", "advisory          : ",
                     "threshold owner   : ",
                     "the Stage-B workbook was bootstrapped"):
        assert evidence in out, f"the startup did not reach {evidence!r}:\n{out[-2000:]}"


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_26_the_dry_run_would_catch_an_uninitialised_script_variable() -> None:
    """SO THE DRY RUN IS NOT VACUOUS. Remove the report-path initialiser - the
    exact defect that ended P8-Z run 2 - and the startup must die on the first
    line it tries to write."""
    if not (PCCM_ROOT / "build" / "phase9_model_check_inspection.json").is_file():
        pytest.skip("Stage A has not been built into pccm/build")
    source = _text()
    # THE ONE THE FIRST WRITTEN LINE READS. `$script:P9Path` is set before the
    # first Write-P9Line in this runner, so removing it proves nothing; the
    # LINE BUFFER is read by that first call, which is the equivalent of what
    # ended P8-Z run 2 - a helper reading script state nobody created.
    broken = source.replace(
        "$script:P9Lines = New-Object System.Collections.ArrayList\n", "", 1)
    assert broken != source
    _code_out, out = _dry_run(broken)
    assert "REACHED THE EXCEL BOUNDARY" not in out, out[-1500:]
    assert "P9Lines" in out, out[-1500:]


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_27_the_cleanup_runs_safely_when_excel_was_never_created() -> None:
    """THE FAILURE PATH, WALKED. Excel creation is made to throw, so the catch
    and the finally run with every handle still $null."""
    if not (PCCM_ROOT / "build" / "phase9_model_check_inspection.json").is_file():
        pytest.skip("Stage A has not been built into pccm/build")
    source = _text()
    broken = source.replace(
        "    $excel = New-Object -ComObject Excel.Application",
        "    throw 'DRY RUN: Excel could not be created'", 1)
    assert broken != source
    broken = broken.replace(
        "& $bootstrap -BuildDir $tempRoot -Force\n$bootstrapExit = $LASTEXITCODE",
        "Set-Content -LiteralPath $stageBPath -Value 'dry run' -Encoding UTF8\n"
        "$bootstrapExit = 0")
    broken = broken.replace("if ($revision.Dirty.Count -gt 0) {",
                            "if ($false -and $revision.Dirty.Count -gt 0) {")
    scratch = WINDOWS / "__p9_cleanup_tmp.ps1"
    scratch.write_text(broken, encoding="utf-8")
    try:
        done = subprocess.run([PWSH, "-NoProfile", "-File", str(scratch)],
                              cwd="/tmp", capture_output=True, text=True, timeout=300)
    finally:
        scratch.unlink(missing_ok=True)
    out = done.stdout + done.stderr
    assert "DRY RUN: Excel could not be created" in out, out[-2500:]
    # THE CLEANUP COMPLETED AND THE VERDICT WAS STILL WRITTEN.
    assert "COM LIFECYCLE" in out, f"the finally block did not finish:\n{out[-2500:]}"
    assert "P9-1 FAIL" in out, out[-2500:]
    assert done.returncode == 1, done.returncode


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
