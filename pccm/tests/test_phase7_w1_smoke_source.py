#!/usr/bin/env python3
"""P7-7: the MINIMAL W1 runner, proved on Linux before a Windows session.

WHY THERE IS A SECOND W1 IMPLEMENTATION AT ALL. Three execution-layer defects
came out of `phase7_acceptance_scenarios.ps1` in a row - a PowerShell 6+ path
construction, a transitive helper that was never in scope, and a procedure
detector that reported eight false absences while the very same procedures
answered Application.Run seconds later - and the third of those also stranded
the owned Excel process. The harness is frozen as history; the Windows
execution authority for W1 is now `phase7_w1_smoke.ps1`, which is small enough
to read in one sitting and depends on one definition-only file.

WHAT THIS FILE CAN AND CANNOT PROVE. There is no PowerShell and no Excel here,
so nothing below claims the runner RAN. What it proves is what a reader would
otherwise take on trust: that the runner's scope really is W1 and nothing else,
that the detector that failed is gone and its replacement discriminates on the
real module sources, that the command surface it looks for is the authorised
one and matches the contract projection, and that every COM object it acquires
has a named release with the shutdown ordering the accepted lifecycle policy
requires.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase7_acceptance_harness_source as accepted  # noqa: E402

WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
RUNNER = WINDOWS / "phase7_w1_smoke.ps1"
LIFECYCLE = WINDOWS / "com_lifecycle.ps1"
FROZEN = WINDOWS / "phase7_acceptance_scenarios.ps1"
VBA = PCCM_ROOT / "src" / "vba"
BUILD = PCCM_ROOT / "build"

# THE AUTHORISED W1 SURFACE, restated here as the independent authority the
# runner's own table is checked against. Eight procedures, four modules.
REQUIRED_SURFACE = {
    "modSimAnnualRun": ("PCCM_RunAnnualStochastic",),
    "modSimAnnualStore": ("PCCM_AnnualDistributionState", "PCCM_AnnualProfileState",
                          "PCCM_AnnualProfilePx", "PCCM_AnnualYearCount"),
    "modSimPostReport": ("PCCM_RunSensitivity",),
    "modSimReport": ("PCCM_RunSimulation", "PCCM_SimulationStatus"),
}

# The frozen harness at the commit the user's W1 was run from. It is evidence,
# not code under maintenance, and this pins it to that.
FROZEN_SHA256 = "9744d9b7c1b4ebbc94ae48db54dd2ffb74af9a554ee43d8da8b1e06efe68c8ed"


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    """The runner with its block comment and its line comments removed.

    Every claim below is about what the runner DOES. Its header explains the two
    defects it exists for and names the very API it refuses to call, so a scan
    that read the prose would convict the file of its own documentation.
    """
    return accepted._ps_code(RUNNER)


def _functions() -> dict[str, str]:
    """name -> body, brace-matched over the COMMENT-STRIPPED source.

    The accepted harness's own reader also removes string literals - it is
    hunting verb-noun tokens, and prose would look like calls to it. Here the
    literals ARE the subject: the declared surface, the detector's pattern and
    the pathspecs are all string literals, so they have to survive.
    """
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
    return out


def _function(name: str) -> str:
    body = _functions().get(name)
    assert body is not None, f"{name} is not defined in the W1 runner"
    return body


def _top_level() -> str:
    return accepted._ps_top_level(RUNNER)


# ===========================================================================
# A. SCOPE: THIS IS W1, AND IT IS SMALL
# ===========================================================================

def test_01_the_runner_exists_and_is_small() -> None:
    """SIZE IS THE POINT. The harness it replaces is 2,123 lines."""
    assert RUNNER.exists(), "the minimal W1 runner is missing"
    lines = _text().splitlines()
    assert len(lines) < 800, (
        f"the W1 runner is {len(lines)} lines; it is meant to be readable in one "
        "sitting, and growing it is how the last one became unrunnable")
    frozen = len(FROZEN.read_text(encoding="utf-8").splitlines())
    assert len(lines) < frozen / 2


def test_02_it_dot_sources_exactly_one_definition_only_file() -> None:
    """THE WHOLE EXTERNAL DEPENDENCY SURFACE, and the reason W1 defect #2 cannot
    happen here: there is no second file whose own helpers might be missing."""
    sourced = re.findall(r"^\. \(Join-Path \$scriptDir '([\w.]+)'\)", _code(), re.M)
    assert sourced == ["com_lifecycle.ps1"], (
        f"the W1 runner dot-sources {sourced}; it may dot-source only the "
        "definition-only lifecycle policy")


def test_03_the_stage_b_bootstrap_is_a_child_process_not_a_dot_source() -> None:
    """`build_stage_b.ps1` has executable top-level code: dot-sourcing it would
    run the bootstrap in this scope. The accepted harnesses invoke it, and so
    does this."""
    code = _code()
    assert "& $bootstrap -BuildDir $tempRoot -Force" in code
    assert ". (Join-Path $scriptDir 'build_stage_b.ps1')" not in code


def test_04_no_w2_to_w8_logic_is_present() -> None:
    """The runner proves eight things. Anything from the behavioural matrix in
    here would be a scenario nobody authorised it to run."""
    code = _code()
    for forbidden in ("Set-Phase5Fixture", "Get-CalcTableRows", "Invoke-Phase5GateBScenarios",
                      "Invoke-Phase6GateBScenarios", "phase7_acceptance_cases",
                      "phase5_gate_b_scenarios.ps1", "phase6_gate_b_scenarios.ps1",
                      "phase4_functional_test.ps1", "Get-Phase5TypedTableBody",
                      "Write-RowObject", "Compare-", "Test-P7SameGrid"):
        assert forbidden not in code, f"{forbidden} is W2-W8 territory"
    for scenario in ("W2", "W3", "W4", "W5", "W6", "W7", "W8"):
        assert f"'{scenario}'" not in code, f"the runner names scenario {scenario}"


def test_05_no_computation_endpoint_is_ever_invoked() -> None:
    """PRESENCE IS THE CLAIM, NOT EXECUTION. `PCCM_RunSimulation`,
    `PCCM_RunSensitivity` and `PCCM_RunAnnualStochastic` appear in the declared
    surface as DATA - names to look for - and must never be run."""
    code = _code()
    runnable = re.findall(r"\$excel\.Run\(([^)]*)\)", code)
    assert runnable, "the runner never calls Application.Run at all"
    for call in runnable:
        for endpoint in ("PCCM_RunSimulation", "PCCM_RunSensitivity",
                         "PCCM_RunAnnualStochastic"):
            assert endpoint not in call, (
                f"{endpoint} is executed by the W1 runner; W1 proves it exists, "
                "it does not run it")
    # The only literal endpoint it runs is the automation guard; everything else
    # is resolved from the contract projection.
    literals = [c for c in runnable if "'" in c]
    assert all("PCCM_AutomationBegin" in c for c in literals), literals


def test_06_the_frozen_harness_is_untouched() -> None:
    """It is evidence of the attempted acceptance matrix, and it stays that way.

    The instruction was to stop patching it, not to tidy it - so this control
    exists to make an edit to it fail here rather than pass unnoticed.
    """
    import hashlib
    digest = hashlib.sha256(FROZEN.read_bytes()).hexdigest()
    assert digest == FROZEN_SHA256, (
        "phase7_acceptance_scenarios.ps1 has been modified; it is frozen as "
        "history and the W1 authority is phase7_w1_smoke.ps1")
    assert "phase7_acceptance_scenarios" not in _code(), (
        "the W1 runner reaches back into the frozen harness; naming it in the "
        "header to say what it replaces is the only mention allowed")


# ===========================================================================
# B. WINDOWS POWERSHELL 5.1
# ===========================================================================

def test_07_the_runner_declares_and_obeys_the_shell_it_targets() -> None:
    text = _text()
    assert "WINDOWS POWERSHELL 5.1" in text
    code = _code()
    assert "$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)" in code
    assert "$repoRoot = Split-Path -Parent $pccmRoot" in code
    # W1 defect #1, refused in this file specifically as well as subtree-wide.
    for number, line in enumerate(code.splitlines(), 1):
        count = accepted._join_path_positional_count(line)
        assert count is None or count <= 2, f"line {number}: {line.strip()}"


def test_08_no_powershell_6_or_7_only_construct_is_used() -> None:
    """The class W1 defect #1 belonged to, applied to the new file."""
    offenders: list[str] = []
    code = _code()
    for label, pattern in accepted.PS51_ONLY_CONSTRUCTS:
        for number, line in enumerate(code.splitlines(), 1):
            if re.search(pattern, line):
                offenders.append(f"{number}: {label}: {line.strip()[:80]}")
    assert not offenders, "\n  ".join(offenders)


def test_09_every_custom_command_the_runner_reaches_is_defined() -> None:
    """W1 defect #2, refused for the new file: transitively, not by name."""
    defined: dict[str, str] = {}
    for path in (RUNNER, LIFECYCLE):
        for name, body in accepted._ps_functions(path).items():
            defined.setdefault(name, body)

    seen: set[str] = set()
    missing: list[tuple[str, str]] = []
    work = [("<runner top level>", _top_level())]
    work += [(name, body) for name, body in accepted._ps_functions(RUNNER).items()]
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
    assert not missing, (
        "custom command(s) called but defined neither in the runner nor in "
        "com_lifecycle.ps1:\n  " +
        "\n  ".join(f"{call}  (reached from {where})" for where, call in sorted(set(missing))))


# ===========================================================================
# C. THE DETECTOR THAT REPLACED THE ONE THAT FAILED
# ===========================================================================

def _detector_pattern(name: str) -> str:
    """The runner's own pattern, REBUILT from its source rather than retyped.

    Retyping it would prove that a copy discriminates and say nothing about the
    detector that will run on Windows. So the `$pattern = ...` expression is
    read out of the function and evaluated the way PowerShell evaluates it:
    each single-quoted literal contributes its text verbatim - a PowerShell
    single-quoted string has no escapes, so `[ \\t]` reaches .NET and Python
    alike as "space or tab" - and `$escaped` contributes the escaped name.
    """
    body = _function("Test-W1ProcedureDeclared")
    match = re.search(r"\$pattern = (.*?)\n(?=\s*return)", body, re.S)
    assert match, "the detector no longer builds a $pattern"
    out = ""
    terms = 0
    for term in re.finditer(r"'([^']*)'|(\$escaped)", match.group(1)):
        terms += 1
        out += re.escape(name) if term.group(2) else term.group(1)
    assert terms >= 2, "the detector's pattern could not be read out of the source"
    assert "$escaped" not in out and out.count("(?im)") == 1
    return out


def _declares(module_text: str, name: str) -> bool:
    return re.search(_detector_pattern(name), module_text) is not None


def test_10_the_detector_is_pure_and_touches_no_com() -> None:
    """A TOTAL FUNCTION IS THE WHOLE CORRECTION. The old detector asked a COM
    object about a line, that call REFUSED a declarations line, and a refusal is
    not an answer - so every caller had to turn an exception into 'absent'. This
    one is a string match: it answers yes or no and cannot raise."""
    body = _function("Test-W1ProcedureDeclared")
    for forbidden in ("$Workbook", "VBProject", "VBComponent", "CodeModule",
                      "$excel", "$Excel", ".Run(", "New-Object", "Release-"):
        assert forbidden not in body, f"the detector touches {forbidden}"
    assert "try" not in body and "catch" not in body, (
        "the detector has a try/catch; a detector that needs one is a detector "
        "that can refuse, which is exactly the defect being corrected")
    assert "[regex]::IsMatch" in body


def test_11_the_refusing_codemodule_apis_are_gone() -> None:
    """THE EXACT ROOT CAUSE, refused by name. `ProcOfLine` raises 'Sub or
    Function not defined' for any line in a module's declarations section, and
    the old detector started at line 1 of a module whose first lines are
    Option Explicit and the Public Consts."""
    code = _code()
    for api in ("ProcOfLine", "ProcStartLine", "ProcCountLines", "ProcBodyLine"):
        assert api not in code, (
            f"the W1 runner calls CodeModule.{api}; those are the APIs that "
            "refuse rather than answer, and they are what broke W1")
    # It reads the module's text once instead.
    assert "$code.Lines(1, $lineCount)" in code
    assert "CountOfLines" in code


def test_12_the_detector_is_given_bare_names_only() -> None:
    """The user's hypothesis, checked and then made structurally impossible.

    The failed checks were LABELLED `modSimAnnualStore.PCCM_...`, but that
    qualification only ever existed in the label: the probe itself was always
    passed a bare name. Here the point is settled in code - no dotted name can
    reach the detector, because the names come from the declared table.
    """
    code = _code()
    calls = re.findall(r"Test-W1ProcedureDeclared[^\n]*-ProcedureName ([^\s]+)", code)
    assert calls, "the detector is never called"
    for argument in calls:
        assert "." not in argument, (
            f"a module-qualified name reaches the detector: {argument}")
    for module, procedures in REQUIRED_SURFACE.items():
        for procedure in procedures:
            assert f"{module}.{procedure}" not in code


def test_13_positive_control_every_required_procedure_is_found() -> None:
    """THE POSITIVE HALF, on the real module sources the workbook is built from."""
    for module, procedures in REQUIRED_SURFACE.items():
        text = (VBA / f"{module}.bas").read_text(encoding="utf-8")
        for procedure in procedures:
            assert _declares(text, procedure), (
                f"the detector cannot find {procedure} in {module}.bas, which "
                "declares it - it would report a false absence exactly as the "
                "old one did")


def test_14_negative_control_the_detector_says_no() -> None:
    """THE NEGATIVE HALF. A detector that answers the same for a name that is
    there and a name that is not has proved nothing, which is precisely what
    the failed W1's detector did: it said 'absent' for all eight."""
    texts = {module: (VBA / f"{module}.bas").read_text(encoding="utf-8")
             for module in REQUIRED_SURFACE}
    sentinel = re.search(r"W1AbsentProcedure = '([\w]+)'", _code()).group(1)
    for module, text in texts.items():
        assert not _declares(text, sentinel), sentinel
        # MODULE-SPECIFIC, not merely name-specific: a real procedure must not
        # be found in the modules it does not live in. The failed detector could
        # not tell one module from another at all.
        for other, procedures in REQUIRED_SURFACE.items():
            if other == module:
                continue
            for procedure in procedures:
                assert not _declares(text, procedure), (
                    f"{procedure} is reported present in {module}, but it lives "
                    f"in {other}")
    # AND THE TWO HALVES DISAGREE, which is the property being proved.
    store = texts["modSimAnnualStore"]
    assert _declares(store, "PCCM_AnnualYearCount")
    assert not _declares(store, sentinel)


def test_15_the_detector_refuses_the_near_misses() -> None:
    """A prefix, a comment, a call site and an End Sub are not declarations."""
    text = (VBA / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    assert not _declares(text, "PCCM_AnnualYear"), (
        "a prefix of a real procedure name is reported present; the trailing "
        "identifier guard is not doing its job")
    synthetic = "\n".join([
        "Option Explicit",
        "' Public Function PCCM_Commented() As String",
        "Rem Public Sub PCCM_Remmed()",
        "Public Sub PCCM_Real()",
        "    Call PCCM_Elsewhere",
        "End Sub",
    ])
    assert _declares(synthetic, "PCCM_Real")
    assert not _declares(synthetic, "PCCM_Commented")
    assert not _declares(synthetic, "PCCM_Remmed")
    assert not _declares(synthetic, "PCCM_Elsewhere")


def test_16_the_runner_runs_both_controls_in_the_windows_session() -> None:
    """A Linux proof of the pattern is not a Windows proof of the session.

    So the runner asks the same detector, in the same session, about names that
    ARE there and names that are NOT, and refuses a run in which those two
    answers ever agree.
    """
    code = _code()
    assert "$script:W1AbsentProcedure" in code
    assert "is NOT declared in" in code, "the cross-module negative control is missing"
    assert "the detector does NOT find" in code, "the sentinel control is missing"
    assert "the procedure detector discriminates" in code, (
        "the runner does not require its detector to have answered both ways")
    assert "($positives -gt 0) -and ($negatives -gt 0)" in code


# ===========================================================================
# D. THE DECLARED SURFACE
# ===========================================================================

def _declared_surface() -> dict[str, tuple[str, ...]]:
    body = _function("Get-W1RequiredSurface")
    out: dict[str, tuple[str, ...]] = {}
    for match in re.finditer(r"Module = '(\w+)'\s*Procedures = @\(([^)]*)\)", body, re.S):
        out[match.group(1)] = tuple(re.findall(r"'([\w]+)'", match.group(2)))
    return out


def test_17_the_declared_surface_is_the_authorised_one() -> None:
    assert _declared_surface() == REQUIRED_SURFACE


def test_18_every_declared_procedure_is_public_in_that_module_and_no_other() -> None:
    """The table describes the source tree, and this is what says so."""
    sources = {path.stem: path.read_text(encoding="utf-8") for path in VBA.glob("*.bas")}
    for module, procedures in REQUIRED_SURFACE.items():
        for procedure in procedures:
            assert re.search(rf"^Public (?:Sub|Function) {procedure}\b", sources[module], re.M), (
                f"{procedure} is not a Public procedure of {module}")
            for other, text in sources.items():
                if other == module:
                    continue
                assert not re.search(rf"^\s*(?:Public |Private |Friend )?(?:Sub|Function) "
                                     rf"{procedure}\b", text, re.M), (
                    f"{procedure} is also declared in {other}")


def test_19_the_contract_projection_is_covered_and_checked_at_run_time() -> None:
    """A contract that renames a command must not leave the runner looking for
    the old name and reporting a pass, so the runner cross-checks its own table
    against the projection before it uses it."""
    projection = json.loads(
        (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8"))
    surface = projection["command_surface"]
    projected = [surface["annual_endpoint"]] + list(surface["handoff_accessors"])
    declared = [name for names in REQUIRED_SURFACE.values() for name in names]
    assert not [name for name in projected if name not in declared]
    code = _code()
    assert "$p7.command_surface.annual_endpoint" in code
    assert "the declared W1 surface covers every contract-projected command" in code


def test_20_the_safe_accessors_are_the_four_and_their_unrun_answers() -> None:
    """Only the read-only four are called, and the expected answers come from
    the projection rather than from strings typed into the runner."""
    code = _code()
    assert "$p7.handoff.distribution_states[0]" in code
    assert "$p7.handoff.profile_states[0]" in code
    assert "'NOT PRODUCED'" not in code, (
        "the unrun state is typed into the runner; it is the contract's to say")
    projection = json.loads(
        (BUILD / "phase7_acceptance_inspection.json").read_text(encoding="utf-8"))
    assert projection["handoff"]["distribution_states"][0] == "NOT PRODUCED"
    assert projection["handoff"]["profile_states"][0] == "NOT PRODUCED"
    assert projection["command_surface"]["handoff_accessors"] == [
        "PCCM_AnnualDistributionState", "PCCM_AnnualProfileState",
        "PCCM_AnnualProfilePx", "PCCM_AnnualYearCount"]
    # Blank and zero are the other two, and they are checked as such.
    assert "'BLANK'" in code and "'ZERO'" in code
    assert "[double]$value -eq 0" in code


# ===========================================================================
# E. COM LIFECYCLE AND SHUTDOWN
# ===========================================================================

def test_21_every_acquired_com_object_has_a_named_release() -> None:
    """The seven the authorisation names, each released by the accepted policy."""
    code = _code()
    for label in ("'CodeModule'", "'VBComponent'", "'VBComponents'", "'VBProject'"):
        assert f"Release-Transient" in code and label in code, label
    for label in ("'Workbook'", "'Workbooks'", "'Excel.Application'"):
        assert f"Invoke-NamedRelease $rel" in code and label in code, label


def test_22_no_com_object_escapes_the_project_read() -> None:
    """PLAIN DATA CROSSES THE BOUNDARY. A returned VBComponent would be a
    reference nothing later in the run knows it is holding."""
    body = _function("Read-W1VbaProject")
    assert body.count("Release-Transient") == 4, (
        "the four VBE objects are CodeModule, VBComponent, VBComponents and "
        "VBProject; each needs its own named release")
    returned = re.search(r"return \$result", body)
    assert returned, "the project read returns something other than its result record"
    # The record is built from names, strings and a boolean only.
    assert "Loaded" in body and "Texts" in body and "TrustRefused" in body
    assert "$result.Texts[$name] = $text" in body
    for leak in ("$result.Component", "$result.CodeModule", "$result.Project"):
        assert leak not in body
    # AND THE PER-COMPONENT RELEASE IS IN A finally, so a module that cannot be
    # read still releases the two objects opened for it.
    assert "} finally {" in body


def test_23_the_error_collection_is_cleared_before_the_collect() -> None:
    """THE EXACT REASON PID 27384 SURVIVED Application.Quit.

    A failed COM call leaves an ErrorRecord that still references the object the
    call was made on; $Error keeps up to 256 of them for the life of the
    session; and a rooted RCW is one the collects cannot reclaim, so Excel's
    reference count never reaches zero.
    """
    code = _code()
    assert "$Error.Clear()" in code
    clear_at = code.index("$Error.Clear()")
    collect_at = code.index("[System.GC]::Collect()")
    wait_at = code.index("Wait-ExcelExit")
    assert clear_at < collect_at < wait_at, (
        "the order must be clear, collect, wait: clearing after the collect "
        "leaves the RCW rooted for the collect that mattered")


def test_24_the_shutdown_is_the_accepted_path_on_every_route() -> None:
    code = _code()
    session = code[code.index("$excel = New-Object -ComObject Excel.Application"):]
    shutdown = session[session.index("} finally {"):]
    for required in ("$wb.Close($false)", "$excel.Quit()", "Wait-ExcelExit",
                     "Invoke-EmergencyExcelCleanup", "EXCEL SHUTDOWN"):
        assert required in shutdown, required
    assert shutdown.count("finally") >= 2, (
        "the shutdown has no inner finally; a release that throws would skip "
        "the wait and orphan the process")
    assert shutdown.count("EXCEL SHUTDOWN:") >= 2, (
        "the report must say what became of the owned process on the normal "
        "path too, not only when cleanup was needed")


def test_25_emergency_cleanup_can_never_produce_a_pass() -> None:
    """It stays as a safety net - a failed run must not leave a process behind -
    but a run that needed it FAILS, and the verdict is what says so."""
    code = _code()
    assert "$emergencyRequired = $true" in code
    assert "'no emergency cleanup was required' (-not $emergencyRequired)" in code
    assert "'the owned Excel process exited naturally' $naturalExit" in code
    assert "'every transient COM release succeeded'" in code
    # The verdict is a plain conjunction over the checks, so those three cannot
    # be recorded and then ignored.
    assert "$failed = @($script:W1Checks | Where-Object { -not $_.Ok })" in code
    assert "$ok = (($failed.Count -eq 0)" in code
    assert "if ($ok) { exit 0 } else { exit 1 }" in code


def test_26_no_ledger_property_is_invented() -> None:
    """A DEFECT CLASS THE FROZEN HARNESS ACTUALLY HAS.

    It assigns `$rel.Quit`, and the ledger's property is `QuitCalled`; under
    Set-StrictMode 2.0 that assignment throws and is caught as an
    'Application.Quit' failure, so a successful Quit is reported as a failed
    one. The frozen file is not edited - it is evidence - but the new runner
    must not repeat it, and this is what checks.
    """
    ledger = re.search(r"function New-ReleaseLedger \{(.*?)^\}",
                       LIFECYCLE.read_text(encoding="utf-8"), re.M | re.S).group(1)
    properties = set(re.findall(r"^\s{8}(\w+)\s*=", ledger, re.M))
    assert {"WorkbookClosed", "QuitCalled", "NaturalExit", "EmergencyRequired"} <= properties
    used = set(re.findall(r"\$rel\.(\w+)\s*=", _code()))
    assert used <= properties, (
        f"the W1 runner assigns ledger properties that do not exist: "
        f"{sorted(used - properties)}")


def test_27_the_owned_process_is_the_only_one_ever_touched() -> None:
    """No security setting, no registry, no process this run did not create."""
    code = _code()
    assert "Get-PreExistingExcelPids" in code
    assert "Get-ExcelIdentity" in code
    assert "Stop-Process" not in code, (
        "force-stopping belongs to Invoke-EmergencyExcelCleanup, which verifies "
        "the identity first")
    for forbidden in ("AccessVBOM", "Registry", "HKCU", "HKLM", "TrustedLocation",
                      "VBAWarnings", "Set-ItemProperty"):
        assert forbidden not in code, forbidden


def test_28_the_real_build_is_copied_never_opened_and_never_saved() -> None:
    code = _code()
    assert "$tempRoot = Join-Path ([System.IO.Path]::GetTempPath())" in code
    assert "$workbooks.Open($stageBPath)" in code
    assert "$stageBPath = Join-Path $tempRoot" in code
    for forbidden in (".Save()", ".SaveAs(", ".SaveCopyAs("):
        assert forbidden not in code, f"the W1 runner {forbidden}"
    assert "$wb.Close($false)" in code


# ===========================================================================
# F. THE CANDIDATE IDENTITY AND THE REPORT
# ===========================================================================

def test_29_the_tree_is_proved_clean_before_excel_is_started() -> None:
    code = _code()
    revision = _function("Get-W1SourceRevision")
    for pathspec in ("'pccm/src', 'pccm/spec', 'pccm/builder'",):
        assert pathspec in revision
    assert "rev-parse HEAD" in revision
    refusal_at = code.index("REFUSED, BEFORE EXCEL WAS STARTED.")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert refusal_at < excel_at
    assert "exit 1" in code[refusal_at:excel_at]


def test_30_the_report_is_written_through_on_every_line() -> None:
    """A run that is stopped, or one that hangs, still leaves what it had."""
    body = _function("Write-W1Line")
    assert "Set-Content -LiteralPath $script:W1Path" in body
    assert "$script:W1Lines" in body


def test_31_the_eight_proof_points_are_each_a_recorded_check() -> None:
    """The authorisation lists eight; a proof point with no check is a claim."""
    code = _code()
    for claim in ("the Stage-B workbook was generated from the current Stage-A build",
                  "manifest-declared modules are loaded in the project",
                  "the complete VBAProject compiles in real Excel",
                  "is declared in",
                  "is callable through Application.Run",
                  "answers the unrun state",
                  "the owned Excel process exited naturally",
                  "no emergency cleanup was required"):
        assert claim in code, claim
    # And the clean-tree proof point refuses the run outright rather than
    # recording a check, which is stronger.
    assert "could not be attributed to a source revision" in _text()


def test_32_a_compile_failure_stops_the_run() -> None:
    """Nothing observed after a failed compile is evidence about anything."""
    code = _code()
    compile_at = code.index("the complete VBAProject compiles in real Excel")
    surface_at = code.index("is declared in ")
    assert compile_at < surface_at
    assert "throw ('the VBAProject does not compile: '" in code
    # THE AUTOMATION CALL IS THE FIRST Application.Run, so it is where a compile
    # failure actually surfaces; it feeds the compile check rather than escaping
    # as a fatal with no compile evidence recorded.
    assert "$compileFailure = $automationProblem" in code
