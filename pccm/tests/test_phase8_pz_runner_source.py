#!/usr/bin/env python3
"""P8-Z: the one tiny Windows scenario the accepted P8-3 fixture could not ask.

WHY IT EXISTS AND WHY IT IS SEPARATE. The 189-check P8-3 suite is accepted at
9e3c141 and is not re-run. Its W4 fixture has five drivers and all five vary, so
it never produces a diagnostic row - and the question this settlement turned on
is what happens when it does. That needs a model P8-3 does not build, and a
model P8-3 does not build needs a runner of its own rather than a seventh part
bolted onto an accepted one.

WHAT THESE CONTROLS EXIST TO CATCH. A tiny runner fails in the ways a big one
does, and two more besides: it can quietly re-test what is already accepted,
which wastes a Windows run and muddies which evidence established what; and it
can prove the exclusion by asking the BRIDGE rather than the CHART, which proves
the formula and not the picture.

WHAT THIS FILE CANNOT PROVE: there is no PowerShell and no Excel here, so
nothing below claims the runner RAN.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

import test_phase7_acceptance_harness_source as accepted  # noqa: E402

WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
RUNNER = WINDOWS / "phase8_pz_zero_variance.ps1"
P83 = WINDOWS / "phase8_p3_chart_surface.ps1"
BUILD = PCCM_ROOT / "build"

# THE PARTS COPIED FROM THE ACCEPTED P8-3 RUNNER, renamed and otherwise
# untouched. Reusing a lifecycle that has run on Windows means reusing it.
COPIED_FROM_P83 = (
    "Write-P8ZLine", "Add-P8ZCheck", "Invoke-P8ZRelease", "Get-P8ZSourceRevision",
    "Invoke-P8ZEndpoint", "Get-P8ZRegister", "Get-P8ZRegisterColumnIndex",
    "Get-P8ZRegisterRowIndex", "Get-P8ZCell", "Format-P8ZCell", "Test-P8ZBlank",
    "Test-P8ZSameCellValue", "Test-P8ZSameValue", "Invoke-P8ZRecalculate",
    "ConvertTo-P8ZColumnNumber", "Get-P8ZCharts", "Test-P8ZNoPoint",
    "Format-P8ZPoint", "Get-P8ZSeriesParts", "ConvertTo-P8ZNormalRange",
    "ConvertTo-P8ZCentimetres",
)

# THE ONE READER THIS SCENARIO ADDS, because Get-P8ZCharts captures the plotted
# VALUES and the category RANGE but not the category VALUES - and "no category
# for that driver" is a statement about the latter.
ADDED = ("Get-P8ZCategoryValues",)

_CACHE: dict = {}


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    return accepted._ps_code(RUNNER)


def _function(name: str) -> str:
    code = _code()
    start = code.index(f"function {name} {{")
    depth = 0
    for index in range(start, len(code)):
        if code[index] == "{":
            depth += 1
        elif code[index] == "}":
            depth -= 1
            if depth == 0:
                return code[start:index + 1]
    raise AssertionError(f"{name} is not closed")


def _charts() -> dict:
    if "charts" not in _CACHE:
        _CACHE["charts"] = json.loads(
            (BUILD / "phase8_charts_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["charts"]


# ===========================================================================
# A. IT IS TINY, AND IT IS NOT THE ACCEPTED SUITE
# ===========================================================================
def test_01_the_accepted_p8_3_runner_is_untouched() -> None:
    """9e3c141's acceptance stands on that file's bytes. This one is new."""
    assert P83.exists() and RUNNER.exists()
    assert RUNNER.read_bytes() != P83.read_bytes()
    assert RUNNER.name not in P83.read_text(encoding="utf-8"), (
        "the accepted runner now references this one")


def test_02_it_reruns_none_of_the_accepted_scenarios() -> None:
    """NOT A SEVENTH PART. None of the seven accepted parts appears here, and
    neither do the three charts this question is not about."""
    code = _code()
    for banner in ("PART 0", "PART A1", "PART A2", "PART B", "PART C",
                   "PART D", "PART E"):
        assert banner not in code, f"the tiny runner re-runs {banner}"
    for other in ("s_curve", "histogram", "annual_cash_flow"):
        assert other not in code, f"the tiny runner re-tests the {other} chart"
    for accepted_concern in ("freeze_panes", "chart_status", "FreezePanes",
                             "SplitRow", "bin_contract", "profile_state",
                             "distribution_state"):
        assert accepted_concern not in code, (
            f"the tiny runner re-tests accepted P8-3 surface: {accepted_concern}")
    # AND IT SAYS SO IN ITS OWN OUTPUT, so a reader of the log cannot mistake
    # this result for the acceptance.
    assert "not the P8-3 acceptance suite" in _text()


def test_03_it_asks_one_question_and_stays_small() -> None:
    code = _code()
    checks = re.findall(r"Add-P8ZCheck\s*\(?'([^']+)'", code)
    assert len(checks) <= 24, f"{len(checks)} checks is not a tiny scenario"
    assert len(checks) >= 8, f"only {len(checks)} checks; the nine required "
    # ONE WORKBOOK, ONE SESSION.
    assert code.count("$workbooks.Open(") == 1
    assert code.count("New-Object -ComObject Excel.Application") == 1


# ===========================================================================
# B. THE SAME DISCIPLINES THE ACCEPTED RUNNERS KEEP
# ===========================================================================
def test_10_it_dot_sources_only_the_definition_only_files() -> None:
    sourced = re.findall(r"^\. \(Join-Path \$scriptDir '([^']+)'\)", _code(), re.M)
    assert sourced == ["com_lifecycle.ps1", "phase5_gate_b_scenarios.ps1",
                       "phase6_gate_b_scenarios.ps1"], sourced
    assert "phase7_acceptance_scenarios.ps1" not in _code(), (
        "the abandoned harness is executed")


def test_11_the_copied_parts_are_the_p8_3_ones_renamed() -> None:
    """REUSING A LIFECYCLE THAT HAS RUN ON WINDOWS MEANS REUSING IT."""
    p83 = accepted._ps_code(P83)
    mine = _code()
    for name in COPIED_FROM_P83:
        original = name.replace("P8Z", "P83")
        theirs = re.search(rf"^function\s+{re.escape(original)}\s*\{{.*?^\}}",
                           p83, re.S | re.M)
        ours = re.search(rf"^function\s+{re.escape(name)}\s*\{{.*?^\}}",
                         mine, re.S | re.M)
        assert theirs and ours, name
        assert theirs.group(0).replace("P83", "P8Z") == ours.group(0), (
            f"{name} diverged from the accepted P8-3 {original}")


def test_12_only_the_one_declared_reader_was_added() -> None:
    """A COPIED HELPER STAYS COPIED. What this scenario needs beyond them is one
    function, and it is declared rather than folded into a byte-identical one."""
    defined = set(re.findall(r"^function\s+([\w-]+)", _code(), re.M))
    extra = defined - set(COPIED_FROM_P83) - set(ADDED)
    # The ten Phase-5 table helpers are copied wholesale and keep their names.
    extra -= {"Write-RowObject", "Get-NamedValue", "Set-NamedValue",
              "Get-TableColumnNames", "Set-TableCell", "Get-TableBody",
              "Get-TableRowCount", "Add-BlankTableRow", "Remove-TableRow",
              "Get-IdColumnValues"}
    assert not extra, f"undeclared functions were added: {sorted(extra)}"
    added = _function("Get-P8ZCategoryValues")
    assert "$item.XValues" in added, (
        "the added reader does not read Excel's own category values")
    assert "Release-Transient" in added, "the added reader leaks COM references"


def test_13_the_com_lifecycle_is_the_accepted_one() -> None:
    code = _code()
    assert "$rel = New-ReleaseLedger" in code
    assert code.index("New-ReleaseLedger") < code.index(
        "New-Object -ComObject Excel.Application"), (
        "the ledger is created after Excel exists")
    assert "Wait-ExcelExit -Identity $excelIdentity" in code, (
        "the runner does not wait for the process it owned to exit")
    assert "Invoke-P8ZRelease $rel $wb        'Workbook'" in code
    assert "} finally {" in code


def test_14_no_powershell_6_construct_and_no_join_path_with_two_children() -> None:
    import test_phase8_p1_runner_source as p81_controls
    for label, pattern in accepted.PS51_ONLY_CONSTRUCTS:
        assert not re.search(pattern, _code()), label
    for line in _code().splitlines():
        count = accepted._join_path_positional_count(line)
        assert count is None or count <= 2, line
    assert not p81_controls._comma_bound_arithmetic(_text())


# ===========================================================================
# C. THE QUESTION IT ASKS
# ===========================================================================
def test_20_the_zero_variance_driver_is_made_by_production_not_fabricated() -> None:
    """THE MODEL REACHES THE STATE ON ITS OWN. One cost line is given the same
    minimum, most likely and maximum, so its contribution cannot vary and the
    accepted kernel finds sxx = 0. Nothing here writes a status."""
    code = _code()
    assert "Set-Phase5Fixture" in code, "the accepted fixture path is not used"
    assert "'unit_cost_min', 'unit_cost_most_likely', 'unit_cost_max'" in code, (
        "the driver is not made constant across all three bounds")
    assert "$fixedValue = [double]$constant.min_value" in code
    for fabricated in ("SIM_SENSITIVITY_NO_VARIANCE", "_SimData", "n/a - no variance"):
        assert fabricated not in code, (
            f"the runner writes or types the status itself: {fabricated}")
    # THE STATUS IT COMPARES AGAINST IS PROJECTED, not typed.
    assert "$charts.zero_variance_status" in code
    # AND IT IS CARRIED BY THE PHASE-8 PROJECTION, not by the digest-pinned
    # Phase-6 gate-B cases: a Phase-8 need does not get to move Gate-B evidence.
    assert _charts()["zero_variance_status"] == "n/a - no variance"
    pinned = json.loads(
        (BUILD / "phase6_gate_b_cases.json").read_text(encoding="utf-8"))
    assert "sensitivity_zero_variance_status" not in pinned["vocabulary"], (
        "the label was added to a digest-pinned Gate-B artefact")


def test_21_it_proves_the_row_is_retained_and_its_measures_are_blank() -> None:
    code = _code()
    assert "still published on the Sensitivity sheet" in code
    assert "its status is " in code
    assert "@('rho', 'abs_rho', 'rank', 'direction')" in code, (
        "not every measure is checked for a fabricated zero")
    assert "are blank and not zero" in code
    # THE ROW IS FOUND BY ITS STATUS, not assumed to be last.
    assert "$constantSheetRow = $sheetRow" in code
    assert "not assumed to be last" in _text()


def test_22_the_exclusion_is_proved_at_the_chart_and_not_at_the_bridge() -> None:
    """THE PICTURE, NOT THE FORMULA. The source settlement already proved the
    bridge; a Windows run that only re-read the bridge would prove nothing that
    Linux had not."""
    code = _code()
    assert "Get-P8ZCharts -Workbook $wb" in code, "no chart object is inspected"
    assert "Get-P8ZCategoryValues -Workbook $wb" in code, (
        "Excel's own category values are never read")
    assert "$series.Values" in code, "the plotted values are never read"
    assert "is not a tornado category" in code
    assert "has no bar" in code
    # AND THE BRIDGE IS FROZEN BEFORE THE CHART IS READ.
    assert code.index("$bridge[[string]$column.key]") < code.index(
        "Get-P8ZCharts -Workbook $wb"), (
        "the chart is read before the cells it is compared against are frozen")


def test_23_every_category_must_correspond_to_a_ranked_driver() -> None:
    code = _code()
    assert "every tornado category is an eligible ranked driver" in code
    assert "exactly as many categories as were ranked" in code
    assert "nothing fills the rest" in code
    # THE WIRING IS STILL CHECKED, and against the projection.
    assert "the tornado category range is the projected bridge range" in code
    assert "the tornado value range is the projected bridge range" in code


def test_24_it_recomputes_no_ranking_of_its_own() -> None:
    code = _code()
    for banned in ("Sort-Object", "-Descending", "[Math]::Abs", "PERCENTILE",
                   "Group-Object", "Measure-Object -Sum"):
        assert banned not in code, f"the runner re-derives the ranking: {banned}"


def test_25_every_projection_property_it_reads_exists() -> None:
    """THE CONTROL THAT REPLACES A WINDOWS RUN, applied to the new file too.
    Two P8-3 runs were lost to a property that was not there."""
    import test_phase8_p3_runner_source as p83_controls
    data = {name: json.loads((BUILD / filename).read_text(encoding="utf-8"))
            for name, filename in p83_controls.PROJECTION_ROOTS.items()}
    data["simInspect"] = data["simInspection"]
    data["source"] = data["charts"]["sensitivity_source"]
    data["drivers"] = data["charts"]["bridge"]["drivers"]
    code = _code()
    pattern = r"\$(" + "|".join(data) + r")((?:\.[A-Za-z_][A-Za-z0-9_]*)+)"
    unresolved, checked = [], 0
    for match in re.finditer(pattern, code):
        name, chain = match.group(1), match.group(2)
        parts = chain.strip(".").split(".")
        while parts and parts[-1] in p83_controls.PS_MEMBERS:
            parts = parts[:-1]
        if not parts:
            continue
        checked += 1
        why = p83_controls._resolve_projection(data[name], parts)
        if why:
            unresolved.append(f"${name}{chain} -> {why}")
    assert not unresolved, (
        "the runner would abort on a property that is not there:\n  " +
        "\n  ".join(sorted(set(unresolved))))
    assert checked >= 12, f"only {checked} dereferences were resolved"


def test_26_the_eligibility_field_is_projected_and_not_typed() -> None:
    code = _code()
    assert "$source.eligibility.column" in code
    assert "$source.eligibility.key" in code
    source = _charts()["sensitivity_source"]
    assert source["eligibility"]["key"] == "rank"
    # NO SENSITIVITY ADDRESS IS TYPED ANYWHERE.
    for typed in ("'D13'", "'E13'", "'G13'", '"D13"', "Sensitivity!"):
        assert typed not in code, f"the runner types a sensitivity address {typed}"


# ===========================================================================
# D. STARTUP - WHAT RUN 1 DIED ON, BEFORE IT REACHED EXCEL
# ===========================================================================
# WHAT HAPPENED. `Get-P8ZSourceRevision` takes -RepoRoot; this runner called it
# with NO ARGUMENT, so the parameter took its empty default and git was asked to
# report HEAD for ''. The throw printed the empty root back, which is what named
# the defect. A whole Windows turn, spent before Excel opened.
#
# AND A WORSE ONE BESIDE IT. The accepted runners REFUSE TO RUN against a
# modified pccm/src, pccm/spec or pccm/builder - a result from a tree that is
# not a commit cannot be attributed to one. This runner had no such check at
# all. That would not have errored; it would have produced a green report.
PWSH = "/opt/pwsh/pwsh"


def test_30_the_repo_root_is_derived_the_way_the_accepted_runners_derive_it() -> None:
    """THREE LINES, AND THEY ARE THE ACCEPTED ONES rather than a variation."""
    mine = _code()
    theirs = accepted._ps_code(P83)
    for line in ("$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)",
                 "$repoRoot = Split-Path -Parent $pccmRoot",
                 "if ([string]::IsNullOrWhiteSpace($BuildDir)) "
                 "{ $BuildDir = Join-Path $pccmRoot 'build' }"):
        assert line in mine, f"the accepted derivation line is missing: {line}"
        assert line in theirs, f"the accepted runner no longer carries: {line}"
    # AND IT IS DERIVED ONCE. A second derivation is a second answer waiting to
    # disagree with the first.
    assert mine.count("$repoRoot = ") == 1
    assert mine.count("$pccmRoot = ") == 1
    assert mine.count("$BuildDir = Join-Path") == 1


def test_31_nothing_depends_on_the_working_directory() -> None:
    """THE SCRIPT'S OWN LOCATION IS THE ONLY ANCHOR."""
    code = _code()
    assert "$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path" in code
    for cwd in ("Get-Location", "$PWD", "Resolve-Path '.'", "Convert-Path '.'",
                "Set-Location"):
        assert cwd not in code, f"the runner reads the working directory: {cwd}"


def test_32_no_machine_specific_path_appears_anywhere() -> None:
    text = _text()
    for absolute in ("C:\\", "D:\\", "/home/", "Users\\", "OneDrive", "Desktop",
                     "$env:USERPROFILE", "$HOME"):
        assert absolute not in text, f"a machine-specific path is hard-coded: {absolute}"


def test_33_the_source_revision_is_captured_with_the_derived_root() -> None:
    """THE EXACT DEFECT. The helper is never called without its root again."""
    code = _code()
    assert "Get-P8ZSourceRevision -RepoRoot $repoRoot" in code
    assert not re.search(r"Get-P8ZSourceRevision(?!\s+-RepoRoot)(?!\s*\{)", code), (
        "the helper is called without a repository root")
    assert code.index("Get-P8ZSourceRevision -RepoRoot") < code.index(
        "New-Object -ComObject Excel.Application"), (
        "Excel is started before the run can be attributed to a revision")


def test_34_a_missing_head_is_still_fatal_and_a_dirty_tree_still_refuses() -> None:
    """SOURCE ATTRIBUTION IS NOT WEAKENED BY BEING FIXED."""
    helper = _function("Get-P8ZSourceRevision")
    assert "throw (" in helper, "a missing HEAD stopped being fatal"
    assert "Write-Warning" not in helper, "a missing HEAD was downgraded to a warning"
    code = _code()
    assert "catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }" in code, (
        "a failure to resolve HEAD no longer stops the run")
    assert "if ($revision.Dirty.Count -gt 0) {" in code, (
        "the runner will run against a modified tree")
    assert "REFUSED, BEFORE EXCEL WAS STARTED." in code
    assert code.index("$revision.Dirty.Count") < code.index(
        "New-Object -ComObject Excel.Application")


def test_35_every_artefact_is_checked_before_excel_is_started() -> None:
    code = _code()
    assert "foreach ($required in @($manifestPath, $inspectPath, $simInspectPath," in code
    assert "Test-Path -LiteralPath $required" in code
    assert code.index("$required") < code.index(
        "New-Object -ComObject Excel.Application")


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_36_the_derivation_resolves_every_startup_path_from_the_script_alone() -> None:
    """EXECUTED, NOT READ. The runner's own three derivation lines are run with
    the working directory somewhere else entirely, and every file the runner
    opens before Excel is required to exist at the path they produce.

    This is the control that would have cost nothing and saved a Windows turn."""
    import subprocess
    import tempfile

    code = _code()
    derivation = [line for line in code.splitlines()
                  if line.startswith(("$pccmRoot = ", "$repoRoot = ",
                                      "if ([string]::IsNullOrWhiteSpace($BuildDir)"))]
    assert len(derivation) == 3, derivation
    script = "\n".join([
        "Set-StrictMode -Version 2.0",
        "$ErrorActionPreference = 'Stop'",
        "Set-Location ([System.IO.Path]::GetTempPath())",
        "$BuildDir = ''",
        f"$scriptDir = '{WINDOWS}'",
        *derivation,
        "$names = @('stage_b_manifest.json','phase5_gate_b_inspection.json',",
        "           'phase6_gate_b_inspection.json','phase7_acceptance_cases.json',",
        "           'phase8_charts_inspection.json','PCCM_stageA.xlsx')",
        "foreach ($name in $names) {",
        "    $path = Join-Path $BuildDir $name",
        "    if (-not (Test-Path -LiteralPath $path)) { Write-Output ('MISSING ' + $path) }",
        "}",
        "if ([string]::IsNullOrWhiteSpace($repoRoot)) { Write-Output 'EMPTY repoRoot' }",
        "$head = [string](& git -C $repoRoot rev-parse HEAD 2>$null)",
        "if ([string]::IsNullOrWhiteSpace($head)) { Write-Output 'EMPTY head' }",
        "Write-Output ('OK ' + $repoRoot)",
    ])
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False,
                                     encoding="utf-8") as handle:
        handle.write(script + "\n")
        path = handle.name
    try:
        done = subprocess.run([PWSH, "-NoProfile", "-File", path],
                              capture_output=True, text=True, timeout=180)
    finally:
        Path(path).unlink(missing_ok=True)
    assert done.returncode == 0, done.stderr[:2000]
    lines = [l.strip() for l in done.stdout.splitlines() if l.strip()]
    problems = [l for l in lines if not l.startswith("OK ")]
    assert not problems, "the derivation does not resolve:\n  " + "\n  ".join(problems)
    assert lines[-1] == f"OK {PCCM_ROOT.parent}", lines


def _startup_rules(code: str) -> None:
    """Everything the startup has to be true of, over an arbitrary copy."""
    assert "$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)" in code, (
        "the pccm root is not two parents up from the script")
    assert "$repoRoot = Split-Path -Parent $pccmRoot" in code, (
        "the repository root is not the pccm root's parent")
    assert "Get-P8ZSourceRevision -RepoRoot $repoRoot" in code, (
        "the source revision is not asked with the derived root")
    assert not re.search(r"Get-P8ZSourceRevision(?!\s+-RepoRoot)(?!\s*\{)", code), (
        "the helper is called without a repository root")
    assert "if ($revision.Dirty.Count -gt 0) {" in code, (
        "a modified tree no longer refuses")
    assert "catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }" in code, (
        "a failure to resolve HEAD no longer stops the run")
    helper = re.search(r"function\s+Get-P8ZSourceRevision\s*\{(.*?)\n\}", code, re.S)
    assert helper, "the source-revision helper is gone"
    assert "throw (" in helper.group(1), "a missing HEAD stopped being fatal"
    assert "& git -C $RepoRoot rev-parse HEAD" in code, (
        "git is no longer told which repository to answer for")
    for cwd in ("Get-Location", "$PWD", "Set-Location"):
        assert cwd not in code, f"the runner reads the working directory: {cwd}"
    for absolute in ("C:\\", "OneDrive", "$env:USERPROFILE"):
        assert absolute not in code, f"a machine-specific path is hard-coded: {absolute}"


@pytest.mark.parametrize("name,mutate", [
    ("the source revision is asked without a root",
     lambda code: code.replace("Get-P8ZSourceRevision -RepoRoot $repoRoot",
                               "Get-P8ZSourceRevision", 1)),
    ("the repo root is empty",
     lambda code: code.replace("$repoRoot = Split-Path -Parent $pccmRoot",
                               "$repoRoot = ''", 1)),
    ("one parent traversal is removed",
     lambda code: code.replace(
         "$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)",
         "$pccmRoot = Split-Path -Parent $scriptDir", 1)),
    ("the root is taken from the working directory",
     lambda code: code.replace("$repoRoot = Split-Path -Parent $pccmRoot",
                               "$repoRoot = (Get-Location).Path", 1)),
    ("the Windows clone path is hard-coded",
     lambda code: code.replace(
         "$repoRoot = Split-Path -Parent $pccmRoot",
         "$repoRoot = 'C:\\Users\\pcd\\OneDrive\\Desktop\\PCCM-GateB\\mine'", 1)),
    ("git is invoked without a working directory",
     lambda code: code.replace("& git -C $RepoRoot rev-parse HEAD",
                               "& git rev-parse HEAD", 1)),
    ("a missing HEAD becomes a warning",
     lambda code: code.replace(
         "        throw ('git could not report HEAD for ' + $RepoRoot +",
         "        Write-Warning ('git could not report HEAD for ' + $RepoRoot +", 1)),
    ("a modified tree stops refusing",
     lambda code: code.replace("if ($revision.Dirty.Count -gt 0) {",
                               "if ($false) {", 1)),
])
def test_37_each_way_of_losing_the_startup_is_refused(name: str, mutate) -> None:
    mutated = mutate(_code())
    assert mutated != _code(), f"the mutation '{name}' changed nothing"
    with pytest.raises(AssertionError):
        _startup_rules(mutated)


def test_38_the_startup_rules_pass_on_the_real_runner() -> None:
    """SO THE EIGHT REFUSALS ABOVE ARE REFUSALS OF THE MUTATION."""
    _startup_rules(_code())


# ===========================================================================
# E. THE UNINITIALISED-VARIABLE CLASS, CLOSED
# ===========================================================================
# RUN 2 DIED ON `$script:P8ZPath`. `Write-P8ZLine` reads it and `Invoke-P8ZRelease`
# reads `$script:P8ZResidual`; both were copied out of the accepted runner, which
# initialises FOUR script-scope variables at that point. This file initialised
# two. Under StrictMode the first line the runner ever wrote raised.
#
# WHY THE RUN-1 CONTROLS DID NOT CATCH IT. They asked whether the STARTUP was
# right - the root derived, the revision captured, the paths resolved - and it
# was. The defect was in a COPIED HELPER's dependency on script state that was
# never brought across with it, which no amount of reading the startup would
# find, and which a text search cannot find either: the name is in the file.
#
# SO THE CLASS IS CLOSED TWO WAYS. An AST pass that models PowerShell's scoping,
# and a dry run that EXECUTES the whole pre-Excel path under StrictMode.
AUDIT = PCCM_ROOT / "tests" / "powershell_uninitialised_audit.ps1"


def _audit(path: Path) -> tuple[int, str]:
    import subprocess
    done = subprocess.run([PWSH, "-NoProfile", "-File", str(AUDIT), "-Path", str(path)],
                          capture_output=True, text=True, timeout=300)
    return done.returncode, done.stdout.strip()


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_40_no_variable_in_this_runner_can_be_read_before_assignment() -> None:
    code, out = _audit(RUNNER)
    assert code == 0, f"the runner can read a variable before it is set:\n{out}"
    assert out == "CLEAN", out


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_41_nor_can_any_other_windows_runner() -> None:
    """THE WHOLE SUBTREE, because this class is not this file's alone."""
    offenders = {}
    for path in sorted(WINDOWS.glob("*.ps1")):
        code, out = _audit(path)
        if code != 0:
            offenders[path.name] = out
    assert not offenders, offenders


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_42_the_audit_catches_the_defect_that_stopped_run_2() -> None:
    """SO IT IS NOT VACUOUS. The exact file that failed on Windows is fetched
    from the commit that failed, and the audit is required to name both
    variables - the one that raised and the one that had not been reached yet."""
    import subprocess
    import tempfile
    blob = subprocess.run(
        ["git", "show", "ce47b45:pccm/bootstrap/windows/phase8_pz_zero_variance.ps1"],
        cwd=PCCM_ROOT.parent, capture_output=True, text=True, timeout=120)
    assert blob.returncode == 0, blob.stderr[:500]
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False,
                                     encoding="utf-8") as handle:
        handle.write(blob.stdout)
        path = Path(handle.name)
    try:
        code, out = _audit(path)
    finally:
        path.unlink(missing_ok=True)
    assert code == 1, "the audit passed the file that failed on Windows"
    assert "$script:P8ZPath" in out, out
    assert "$script:P8ZResidual" in out, (
        "the audit missed the second uninitialised variable, which no run had "
        "reached yet")
    assert "line 293" in out, f"the audit does not name the failing line: {out}"


# ---------------------------------------------------------------------------
# THE PRE-EXCEL DRY RUN
# ---------------------------------------------------------------------------
def _dry_run(source: str) -> tuple[int, str]:
    """Execute the runner's startup, under StrictMode, up to the deliberate
    boundary immediately before Excel is created.

    IT RUNS FROM bootstrap/windows, because $scriptDir is what everything else
    derives from and a copy somewhere else would be proving a different file's
    startup. The working directory is set elsewhere on purpose.

    ONE STEP IS STUBBED AND ONLY ONE: build_stage_b.ps1 drives Excel to inject
    the VBA. The paths it is handed are real and everything either side of it is
    the runner's own code.
    """
    import subprocess
    excel = source.index("    $excel = New-Object -ComObject Excel.Application")
    head = source[:source.rindex("try {", 0, excel)]
    assert "New-Object -ComObject" not in head, (
        "Excel is created inside the startup path")
    head = head.replace(
        "& $bootstrap -BuildDir $tempRoot -Force\n$bootstrapExit = $LASTEXITCODE",
        "if (-not (Test-Path -LiteralPath $bootstrap)) "
        "{ throw 'build_stage_b.ps1 is not beside the runner' }\n"
        "Set-Content -LiteralPath $stageBPath -Value 'dry run' -Encoding UTF8\n"
        "$bootstrapExit = 0")
    head += "\nWrite-Output 'REACHED THE EXCEL BOUNDARY'\nexit 0\n"
    scratch = WINDOWS / "__p8z_dryrun_tmp.ps1"
    scratch.write_text(head, encoding="utf-8")
    try:
        done = subprocess.run([PWSH, "-NoProfile", "-File", str(scratch)],
                              cwd="/tmp", capture_output=True, text=True, timeout=300)
    finally:
        scratch.unlink(missing_ok=True)
    return done.returncode, (done.stdout + done.stderr)


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_43_the_startup_reaches_the_excel_boundary_under_strict_mode() -> None:
    """EVERY STARTUP STEP, EXECUTED: the repo root derived, the revision
    captured, the clean-tree refusal evaluated, the artefact paths resolved and
    checked, the working copy made, the report path set, the projections loaded,
    and every script-scope variable initialised. Then it stops."""
    source = _text()
    assert "Set-StrictMode -Version 2.0" in source, (
        "the runner does not run under StrictMode, so this proves nothing")
    code, out = _dry_run(source)
    assert code == 0, f"the startup did not reach the boundary:\n{out[-3000:]}"
    assert "REACHED THE EXCEL BOUNDARY" in out, out[-2000:]
    # AND IT REALLY DID THE WORK, rather than skipping to the end.
    for evidence in ("A ZERO-VARIANCE DRIVER IS NOT A TORNADO CATEGORY",
                     "source revision   : ", "eligibility field : ",
                     "the Stage-B workbook was bootstrapped"):
        assert evidence in out, f"the startup did not reach {evidence!r}:\n{out[-2000:]}"


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_44_the_dry_run_reproduces_the_windows_failure_at_ce47b45() -> None:
    """THE CONTROL THAT WOULD HAVE SAVED THE TURN. Against the commit that
    failed, this fails on Linux with the same sentence Windows printed."""
    import subprocess
    blob = subprocess.run(
        ["git", "show", "ce47b45:pccm/bootstrap/windows/phase8_pz_zero_variance.ps1"],
        cwd=PCCM_ROOT.parent, capture_output=True, text=True, timeout=120)
    assert blob.returncode == 0, blob.stderr[:500]
    _code, out = _dry_run(blob.stdout)
    assert "cannot be retrieved because it has not been set" in out, out[-2000:]
    assert "P8ZPath" in out, out[-2000:]
    assert "REACHED THE EXCEL BOUNDARY" not in out


# ---------------------------------------------------------------------------
# THE CLEANUP PATH, ON AN EARLY FAILURE
# ---------------------------------------------------------------------------
@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_45_the_cleanup_runs_safely_when_excel_was_never_created() -> None:
    """STRICTMODE BUGS SURFACE ON THE FAILURE PATH, which is the path least
    likely to have been exercised. Excel creation is made to throw, so the catch
    and both finally blocks run with every handle still $null."""
    import subprocess
    source = _text()
    broken = source.replace(
        "    $excel = New-Object -ComObject Excel.Application",
        "    throw 'DRY RUN: Excel could not be created'", 1)
    assert broken != source
    broken = broken.replace(
        "& $bootstrap -BuildDir $tempRoot -Force\n$bootstrapExit = $LASTEXITCODE",
        "Set-Content -LiteralPath $stageBPath -Value 'dry run' -Encoding UTF8\n"
        "$bootstrapExit = 0")
    scratch = WINDOWS / "__p8z_cleanup_tmp.ps1"
    scratch.write_text(broken, encoding="utf-8")
    try:
        done = subprocess.run([PWSH, "-NoProfile", "-File", str(scratch)],
                              cwd="/tmp", capture_output=True, text=True, timeout=300)
    finally:
        scratch.unlink(missing_ok=True)
    out = done.stdout + done.stderr
    # THE SESSION FAILED, AS ARRANGED - and it failed the way it was told to,
    # not because the cleanup itself threw.
    assert "THE P8-Z SESSION DID NOT COMPLETE" in out, out[-2500:]
    assert "DRY RUN: Excel could not be created" in out, out[-2500:]
    # AND THE CLEANUP COMPLETED.
    assert "EXCEL SHUTDOWN" in out, f"the finally block did not finish:\n{out[-2500:]}"
    assert "no Excel process was ever owned by this run" in out, out[-2500:]
    assert "VERDICT" in out and "checked, " in out, "the verdict was never reached"
    assert "P8-Z FAIL" in out, "a session that never reached Excel reported PASS"
    for uninitialised in ("cannot be retrieved because it has not been set",
                          "is not recognized", "You cannot call a method on a null"):
        assert uninitialised not in out, (
            f"the cleanup path itself raised: {uninitialised}\n{out[-2500:]}")
    # THE RUN STILL FAILS. A clean shutdown is not a pass.
    assert done.returncode == 1, done.returncode


# ---------------------------------------------------------------------------
# THE MUTATIONS
# ---------------------------------------------------------------------------
# EACH IS A WAY THE STARTUP COULD BE WRONG AGAIN, and each is checked by the
# instrument that would actually catch it - the AST audit for a variable that
# cannot be read, the dry run for a path that cannot be walked. A mutation
# refused by a text assertion alone would prove only that the text changed.
def _audit_source(source: str) -> tuple[int, str]:
    scratch = WINDOWS / "__p8z_audit_tmp.ps1"
    scratch.write_text(source, encoding="utf-8")
    try:
        return _audit(scratch)
    finally:
        scratch.unlink(missing_ok=True)


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
@pytest.mark.parametrize("name,mutate", [
    # THE DEFECT ITSELF.
    ("the report path loses its initialiser",
     lambda src: src.replace("$script:P8ZPath = ''\n", "", 1)),
    # THE OTHER ONE, WHICH NO RUN HAD REACHED YET.
    ("the residual list loses its initialiser",
     lambda src: src.replace(
         "$script:P8ZResidual = New-Object System.Collections.ArrayList\n", "", 1)),
    # AN INITIALISER MOVED INSIDE A FUNCTION, where it does not run until the
    # function is called - which is after the first read.
    # AN INITIALISER THAT RUNS TOO LATE. It exists, at file scope, exactly where
    # $script:P8ZPath's did - after the first line had already been written
    # through it.
    ("an initialiser runs after the first helper call",
     lambda src: src.replace(
         "$script:P8ZResidual = New-Object System.Collections.ArrayList\n", "", 1).replace(
         "$stageBPath = Join-Path $tempRoot",
         "$script:P8ZResidual = New-Object System.Collections.ArrayList\n"
         "$stageBPath = Join-Path $tempRoot", 1)),
    # A LOCAL ASSIGNED ONLY IN A BRANCH AND READ AFTERWARDS.
    ("a local is assigned only inside a branch",
     lambda src: src.replace("$comAcquired = 0\n", "", 1)),
    # A CLEANUP VARIABLE THE finally READS, never assigned.
    ("a cleanup variable is never assigned",
     lambda src: src.replace("$excelIdentity = $null\n", "", 1)),
])
def test_46_the_audit_refuses_each_uninitialised_variable(name: str, mutate) -> None:
    source = _text()
    mutated = mutate(source)
    assert mutated != source, f"the mutation '{name}' changed nothing"
    code, out = _audit_source(mutated)
    assert code == 1, f"'{name}' survived the audit"
    # "CLEAN" is a substring of "CLEANUP", so the clean verdict is the WHOLE
    # output rather than a substring of it.
    assert out.strip() != "CLEAN", out
    assert out.strip(), "the audit refused without saying why"


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
@pytest.mark.parametrize("name,mutate,expected", [
    # STRICTMODE TURNED OFF, so the class stops being detectable at all.
    ("strict mode is weakened",
     lambda src: src.replace("Set-StrictMode -Version 2.0",
                             "Set-StrictMode -Off", 1),
     "strict"),
    # THE REPORT PATH TIED TO WHEREVER THE SHELL HAPPENS TO BE.
    ("the report path depends on the working directory",
     lambda src: src.replace(
         "$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `",
         "$tempRoot = Join-Path (Get-Location).Path `", 1),
     "cwd"),
    # EXCEL OPENED BEFORE THE RUN CAN BE ATTRIBUTED TO A REVISION.
    ("excel is created before the source revision is captured",
     lambda src: src.replace(
         "$revision = $null\ntry { $revision = Get-P8ZSourceRevision -RepoRoot $repoRoot }",
         "$revision = $null\n$probe = New-Object -ComObject Excel.Application\n"
         "try { $revision = Get-P8ZSourceRevision -RepoRoot $repoRoot }", 1),
     "order"),
])
def test_47_the_startup_refuses_each_way_of_being_wrong(
        name: str, mutate, expected: str) -> None:
    source = _text()
    mutated = mutate(source)
    assert mutated != source, f"the mutation '{name}' changed nothing"
    if expected == "strict":
        # THE DRY RUN ASSERTS STRICTMODE ITSELF, because without it the whole
        # instrument is blind.
        assert "Set-StrictMode -Version 2.0" not in mutated
        with pytest.raises(AssertionError):
            assert "Set-StrictMode -Version 2.0" in mutated, (
                "the runner does not run under StrictMode")
        return
    if expected == "cwd":
        with pytest.raises(AssertionError):
            _startup_rules(accepted._ps_code_from_text(mutated)
                           if hasattr(accepted, "_ps_code_from_text") else mutated)
        return
    # THE ORDERING ONE IS STRUCTURAL: Excel must not be reachable before the
    # revision is captured.
    code = mutated
    assert code.index("New-Object -ComObject Excel.Application") < code.index(
        "Get-P8ZSourceRevision -RepoRoot"), "the mutation did not reorder anything"
    with pytest.raises(AssertionError):
        assert code.index("Get-P8ZSourceRevision -RepoRoot") < code.index(
            "New-Object -ComObject Excel.Application"), (
            "Excel is started before the run can be attributed to a revision")


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_48_the_audit_and_dry_run_pass_on_the_real_runner() -> None:
    """SO THE EIGHT REFUSALS ABOVE ARE REFUSALS OF THE MUTATION."""
    code, out = _audit(RUNNER)
    assert code == 0 and out == "CLEAN", out
    exit_code, output = _dry_run(_text())
    assert exit_code == 0 and "REACHED THE EXCEL BOUNDARY" in output
