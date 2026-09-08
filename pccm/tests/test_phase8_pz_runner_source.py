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
    assert len(checks) <= 16, f"{len(checks)} checks is not a tiny scenario"
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
    assert "New-ReleaseLedger" in code
    assert code.index("New-ReleaseLedger") < code.index(
        "New-Object -ComObject Excel.Application"), (
        "the ledger is created after Excel exists")
    assert "Invoke-P8ZRelease -Ledger $ledger" in code
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
