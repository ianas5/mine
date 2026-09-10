#!/usr/bin/env python3
"""P10 / PROBE RUN 4: THE TIMELINE INPUT TYPE DEFECT.

WHAT RUN 4 ESTABLISHED, AND IT IS WORTH KEEPING. On real Windows, with all 14
sheets protected and workbook structure protected, the probe wrote to a
proved-locked table header cell, read it back, restored the original exactly and
found protection still in force. So:

    UserInterfaceOnly=True DOES permit code-driven VALUE writes to a locked cell
    while worksheet protection is active.  -- CONFIRMED

WHAT IT DID NOT ESTABLISH. The probe then failed preparing the timeline inputs:

    System.InvalidCastException: Unable to cast object of type 'System.Double'
    to type 'System.String'.
      at phase10_protection_probe.ps1:175   source: $range.Value2 = $Value

PCCM_ApplyTimeline was NOT INVOKED. Nothing is known about ListObject structural
operations under protection, and the verdict was correctly INCONCLUSIVE.

THE DEFECT IS ONE THIS REPOSITORY HAD ALREADY DIAGNOSED. Phase-5 Runtime Run 4
hit the same exception with the same type pair at
phase5_gate_b_scenarios.ps1:922: PowerShell binds a COM property setter PER CALL
SITE, so a single polymorphic `$x.Value2 = $Value` line cannot carry more than
one CLR type. The accepted helpers all avoid it the same way - one assignment
site per type, each with its own cast. The probe had reimplemented the accepted
Set-NamedValue and dropped both the [double] cast and the ClearContents branch.

These controls keep the reimplementation from coming back, and keep the fix from
being "make everything text" - which would make modTimeline's TryReadDouble
refuse the timeline for a reason no reader could distinguish from protection.

Runs standalone or under pytest.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
PROBE = BOOTSTRAP / "phase10_protection_probe.ps1"
# The accepted, Windows-proven definition the probe must carry a verbatim copy of.
ACCEPTED_HELPER_SOURCE = BOOTSTRAP / "phase4_functional_test.ps1"
# The accepted fixture whose timeline values and setter usage are the authority.
GATE_B_SCENARIOS = BOOTSTRAP / "phase5_gate_b_scenarios.ps1"
TIMING_SCENARIOS = BOOTSTRAP / "phase7_timing_scenarios.ps1"

# The commit this batch started from. Production must be byte-identical to it.
ACCEPTED = "5a2e39a"

_MEMO: dict = {}


def _src(key: str, path: Path) -> str:
    if key not in _MEMO:
        _MEMO[key] = path.read_text(encoding="utf-8")
    return _MEMO[key]


def _probe() -> str:
    return _src("probe", PROBE)


def _probe_code() -> str:
    if "probe_code" not in _MEMO:
        text = re.sub(r"<#.*?#>", "", _probe(), flags=re.S)
        _MEMO["probe_code"] = "\n".join(
            line for line in text.splitlines() if not line.strip().startswith("#"))
    return _MEMO["probe_code"]


def _function(name: str, source: str) -> str:
    """One PowerShell function body, delimited by brace depth."""
    start = source.index(f"function {name} ")
    depth = 0
    seen = False
    for index in range(start, len(source)):
        if source[index] == "{":
            depth += 1
            seen = True
        elif source[index] == "}":
            depth -= 1
            if seen and depth == 0:
                return source[start : index + 1]
    raise AssertionError(f"{name} is not closed")


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=PCCM_ROOT.parent,
                          capture_output=True, text=True, check=True).stdout


def _inspection() -> dict:
    return json.loads((PCCM_ROOT / "build" / "phase5_gate_b_inspection.json")
                      .read_text(encoding="utf-8"))


# ===========================================================================
# A. THE DEFECT ITSELF
# ===========================================================================
def test_01_the_untyped_polymorphic_value2_assignment_is_gone() -> None:
    """REQUIRED CONTROL 1. The exact statement that ended Run 4."""
    code = _probe_code()
    assert "$range.Value2 = $Value" not in code, (
        "the untyped polymorphic assignment that raised the InvalidCastException is back")
    assert "function Set-ProbeNamedValue" not in code, (
        "the reimplemented setter is back")


def test_02_every_value2_assignment_carries_an_explicit_cast() -> None:
    """REQUIRED CONTROL 3. A cast on every branch is not decoration: it is what
    gives that branch its own bound COM call site."""
    code = _probe_code()
    assignments = re.findall(r"\.Value2 = ([^\n}]+)", code)
    assert assignments, code
    for value in assignments:
        assert re.match(r"\[(double|string|bool)\]\$\w+", value.strip()), value.strip()


def test_03_the_defect_is_recorded_where_the_fix_lives() -> None:
    """REQUIRED CONTROL 1. A fix whose reason is not written beside it is one the
    next rewrite deletes."""
    source = _probe()
    assert "InvalidCastException" in source
    assert "Unable to cast object of type 'System.Double'" in source
    assert "phase10_protection_probe.ps1:175" in source
    assert "per call site" in source or "PER CALL SITE" in source
    assert "phase5_gate_b_scenarios.ps1:922" in source, (
        "the prior Windows evidence for this mechanism is not cited")


def test_04_nothing_is_solved_by_stringifying_a_number() -> None:
    """REQUIRED CONTROL 2/3. `[string]$Value` at a numeric site would satisfy the
    binder and then be refused by modTimeline's TryReadDouble - a REFUSED
    endpoint indistinguishable from protection blocking the work."""
    code = _probe_code()
    for banned in ("$range.Value2 = [string]", "$rng.Value2 = [string]$Value",
                   "-Value ([string]"):
        assert banned not in code, f"a numeric input is stringified: {banned}"
    setter = _function("Set-NamedValue", code)
    assert "$rng.Value2 = [double]$Value" in setter


# ===========================================================================
# B. THE ACCEPTED HELPER, REUSED AND PROVED REUSED
# ===========================================================================
def test_10_the_probe_carries_a_verbatim_copy_of_the_accepted_setter() -> None:
    """REQUIRED CONTROL 4. These harnesses are scripts, not modules, so sharing a
    primitive means copying it - and a copy is only worth anything if something
    proves it is one. This is the same rule as the benchmark's test_57."""
    accepted = _function("Set-NamedValue", _src("accepted", ACCEPTED_HELPER_SOURCE))
    assert "$rng.Value2 = [double]$Value" in accepted, (
        "the accepted helper is not the one this control thinks it is")
    assert accepted in _probe(), "the probe's copy has drifted from the accepted definition"


def test_11_the_accepted_helper_itself_is_untouched() -> None:
    """REUSE IS NOT LICENCE TO EDIT. The accepted definition is Phase-4 source and
    its callers are Phase-4 scenarios."""
    for name in ("phase4_functional_test.ps1", "phase5_gate_b_scenarios.ps1",
                 "phase7_timing_scenarios.ps1"):
        diff = _git("diff", "--name-only", ACCEPTED, "--", f"pccm/bootstrap/windows/{name}")
        assert not diff.strip(), f"{name} was edited"


def test_12_null_and_blank_are_handled_explicitly() -> None:
    """REQUIRED CONTROL 5. Dropping this branch is what made Run 2 contradict
    itself, and the probe's own comment already says so."""
    setter = _function("Set-NamedValue", _probe_code())
    assert "if ($null -eq $Value) { $null = $rng.ClearContents() }" in setter
    exact = _function("Set-ProbeCellExact", _probe_code())
    assert "if ($null -eq $Value)        { $null = $Cell.ClearContents() }" in exact


def test_13_an_unsupported_type_is_refused_by_name_not_coerced() -> None:
    exact = _function("Set-ProbeCellExact", _probe_code())
    assert "throw (" in exact
    assert "$Value.GetType().FullName" in exact
    assert "will not write back by coercion" in _function("Set-ProbeCellExact", _probe())


def test_14_each_type_gets_its_own_assignment_site() -> None:
    """REQUIRED CONTROL 10. One line per CLR type is the whole mechanism."""
    exact = _function("Set-ProbeCellExact", _probe_code())
    sites = re.findall(r"\$Cell\.Value2 = \[(\w+)\]", exact)
    assert sites == ["string", "double", "bool"], sites
    assert len(set(sites)) == len(sites), "two types share one assignment site"


# ===========================================================================
# C. THE TIMELINE INPUTS
# ===========================================================================
def test_20_the_three_inputs_are_declared_and_are_the_minimum() -> None:
    """REQUIRED CONTROL 6. PCCM_ApplyTimeline reads a timeline triple; the probe
    sets that triple and nothing else."""
    declared = re.findall(r"@\{ Key = '(\w+)';\s+Value = \[double\](\S+) \}", _probe_code())
    assert declared == [("base_year", "2026"), ("project_start_year", "2027"),
                        ("duration_years", "3")], declared
    assert "discount_rate" not in _probe_code(), "the probe sets more than the command needs"


def test_21_the_values_are_the_accepted_gate_b_shape() -> None:
    """NOT A SECOND FIXTURE CONTRACT. The same triple the accepted Phase-7 timing
    fixture declares, for the reason it already records."""
    timing = _src("timing", TIMING_SCENARIOS)
    assert "base_year = 2026; start_year = 2027; duration = 3" in timing, (
        "the accepted fixture no longer declares the shape this reuses")
    assert "base_year 2026 with start_year 2027 is the accepted shape" in timing


def test_22_the_defined_names_come_from_the_inspection_never_a_literal() -> None:
    code = _probe_code()
    inputs = _inspection()["inputs"]
    for key in ("base_year", "project_start_year", "duration_years"):
        assert f"'{inputs[key]['defined_name']}'" not in code, (
            f"{inputs[key]['defined_name']} is hard-coded")
    assert "-Name 'defined_name'" in code, "the defined name is not read from the inspection"


def test_23_the_inspection_names_and_the_names_production_reads_are_the_same_cells() -> None:
    """THE ALIAS, PROVED RATHER THAN ASSUMED. The probe writes through the inp*
    names the accepted fixture uses; modTimeline reads the nm*_Entered names.
    They resolve to the same three cells today, and nothing but this says so."""
    import openpyxl

    workbook = openpyxl.load_workbook(PCCM_ROOT / "build" / "PCCM_stageA.xlsx")
    constants = (PCCM_ROOT / "build" / "vba" / "modConstants.bas").read_text(encoding="utf-8")
    inputs = _inspection()["inputs"]
    pairs = (("base_year", "NM_BASE_YEAR_ENTERED"),
             ("project_start_year", "NM_PROJECT_START_YEAR_ENTERED"),
             ("duration_years", "NM_DURATION_YEARS_ENTERED"))
    for key, constant in pairs:
        entered = re.search(rf'{constant} As String = "(\w+)"', constants)
        assert entered, constant
        written = workbook.defined_names.get(inputs[key]["defined_name"])
        read = workbook.defined_names.get(entered.group(1))
        assert written is not None and read is not None, key
        assert written.value == read.value, (
            f"{key}: the probe writes {written.value} but production reads {read.value}")


def test_24_the_inputs_are_written_as_double_at_the_call_site_too() -> None:
    """REQUIRED CONTROL 2. Belt and braces: the declaration is [double], the
    parameter is passed [double], and the helper casts [double] again."""
    setter = _function("Set-ProbeTimelineInputs", _probe_code())
    assert "-Value ([double]$entry.Value)" in setter
    assert "Set-NamedValue -Workbook $Workbook" in setter


# ===========================================================================
# D. READBACK BEFORE INVOCATION
# ===========================================================================
def test_30_every_input_is_read_back_before_the_endpoint() -> None:
    """REQUIRED CONTROL 6."""
    check = _function("Test-ProbeTimelineInputs", _probe_code())
    assert "foreach ($entry in @($Inputs))" in check
    assert "Get-ProbeNamedValue -Workbook $Workbook -DefinedName $name" in check


def test_31_the_readback_checks_the_type_and_checks_it_first() -> None:
    """REQUIRED CONTROL 2. A value-only check passes a stringified number."""
    check = _function("Test-ProbeTimelineInputs", _probe_code())
    type_at = check.index("$actual -isnot [double]")
    value_at = check.index("[double]$actual -ne $expected")
    assert type_at < value_at, "the value is compared before the type"
    assert "not System.Double" in check


def test_32_the_reader_does_not_stringify() -> None:
    """THE ACCEPTED Get-NamedValue RETURNS [string]$v BY DESIGN, and that is
    exactly the wrong reader for a type check. The probe keeps a typed one."""
    reader = _function("Get-ProbeNamedValue", _probe_code())
    assert "return $range.Value2" in reader
    # THE VALUE, not the parameter. `[string]$DefinedName` is the name of the
    # range and is meant to be text; what must not be cast is what comes back.
    for banned in ("[string]$range.Value2", "[string]$v", "return [string]"):
        assert banned not in reader, f"the readback reader stringifies: {banned}"
    accepted = _function("Get-NamedValue", _src("accepted", ACCEPTED_HELPER_SOURCE))
    assert "return [string]$v" in accepted, (
        "the accepted reader no longer stringifies, so this distinction is stale")


def test_33_a_precondition_failure_stops_before_the_endpoint() -> None:
    """REQUIRED CONTROL 7."""
    code = _probe_code()
    problems_at = code.index("$inputProblems = @(Test-ProbeTimelineInputs")
    throw_at = code.index("was NOT invoked: ", problems_at)
    invoke_at = code.index("$timeline = Invoke-ProbeEndpoint", problems_at)
    assert problems_at < throw_at < invoke_at, "the readback does not gate the endpoint"
    assert "if (@($inputProblems).Count -gt 0) {" in code


def test_34_the_endpoint_is_still_marked_invoked_only_beside_application_run() -> None:
    """REQUIRED CONTROL 7, and the rule Run 2 earned. $invoked has exactly three
    assignments: the false initialiser and the one beside Application.Run."""
    invoker = _function("Invoke-ProbeEndpoint", _probe_code())
    assignments = re.findall(r"\$invoked = (\$\w+)", invoker)
    assert assignments == ["$false", "$true"], assignments
    true_at = invoker.index("$invoked = $true")
    run_at = invoker.index("$Excel.Run($Endpoint)")
    assert 0 < run_at - true_at < 200, "the flag is not set immediately before the call"
    assert "ENDPOINT INVOKED" in invoker


def test_35_preparing_and_invoking_are_worded_differently() -> None:
    """REQUIRED CONTROL 7 / section 9. A transcript must not be able to imply
    production ran when fixture preparation failed."""
    code = _probe_code()
    assert "SETTING ENDPOINT PRECONDITIONS" in code
    assert "VERIFYING ENDPOINT PRECONDITIONS" in code
    assert code.count("production NOT invoked") == 2
    assert "-Action 'setting the timeline inputs'" not in code, (
        "the Run-4 wording, which sat under stage 'endpoint', is back")


# ===========================================================================
# E. WHAT MUST NOT HAVE MOVED
# ===========================================================================
def test_40_the_locked_cell_control_still_proves_everything_it_did() -> None:
    """REQUIRED CONTROL 11. Run 4's evidence rests on this control, so the
    control's shape is pinned here as well as in its own suite."""
    body = _function("Invoke-ProbeLockedCellControl", _probe_code())
    assert "if (-not $target.Worksheet.ProtectContents) {" in body
    assert "if (-not $cell.Locked) {" in body
    assert "$original = $cell.Value2" in body
    assert "Test-ProbeExactValue -Actual $restored -Expected $original" in body
    assert body.count("Set-ProbeCellExact -Cell $cell") == 2
    # And protection is re-read AFTER the whole thing.
    assert body.rindex("$target.Worksheet.ProtectContents") > body.index("$restored =")


def test_41_run_4_evidence_is_separable_from_the_verdict() -> None:
    """REQUIRED CONTROL 12. The capability answer is a WORD of its own and the
    verdict is a different word; neither is projected onto the other."""
    code = _probe_code()
    results = set(re.findall(r"Result = '(\w+)'", code))
    assert results == {"INCONCLUSIVE", "REFUSED", "SUCCEEDED"}, results
    assert "$controlResult = [string]$control.Result" in code
    assert "$controlWorked" not in code
    assert "That is a SEPARATE capability from permission to perform" in _probe()


def test_42_no_production_vba_or_spec_changed() -> None:
    """REQUIRED CONTROL 13."""
    changed = [line for line in
               _git("diff", "--name-only", ACCEPTED, "--",
                    "pccm/src", "pccm/spec", "pccm/builder").splitlines() if line.strip()]
    assert not changed, changed


def test_43_strict_mode_is_still_on() -> None:
    """REQUIRED CONTROL 14.

    AGAINST THE CODE, NOT THE PROSE. The probe also EXPLAINS StrictMode in a
    comment, and a control that reads the raw file is satisfied by that comment
    while the directive itself is switched off. A mutation proved exactly that.
    """
    code = _probe_code()
    assert "Set-StrictMode -Version 2.0" in code
    assert "Set-StrictMode -Off" not in code


def test_44_clean_com_shutdown_is_still_required() -> None:
    """REQUIRED CONTROL 15.

    WORD-BOUNDED. `Get-TransientFailuresRenamed` contains `Get-TransientFailures`,
    so a substring check calls a renamed - and therefore uncallable - helper
    present. That is the Unprotect/Unprotected false positive, and a mutation
    walked straight through it here.
    """
    code = _probe_code()
    for name in ("Wait-ExcelExit", "Get-TransientFailures",
                 "Invoke-EmergencyExcelCleanup", "Release-Transient"):
        assert re.search(rf"{re.escape(name)}\b(?!-)", code), f"the probe lost {name}"


def test_45_the_probe_still_refuses_to_touch_protection() -> None:
    code = _probe_code()
    for banned in (".Unprotect(", ".Unprotect ", ".Protect(", "'ProtectionRelease'"):
        assert banned not in code, f"the probe now performs {banned}"


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-q"]))
