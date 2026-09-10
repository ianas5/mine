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
def test_20_the_declared_inputs_are_the_minimum_each_command_needs() -> None:
    """REQUIRED CONTROL 6, WIDENED BY ONE ON WINDOWS EVIDENCE.

    Run 6 invoked PCCM_Calculate and it refused for a NON-protection reason:
    "Discount Rate: the value is blank. A blank is not zero." That is
    modCalcResolve.NumericNamedCell doing exactly what it should. The discount
    rate is an ordinary required Setup input, so the probe now supplies it.

    IT IS STILL THE MINIMUM. Every entry names the endpoint that reads it, and
    nothing is set for a command that does not. With zero drivers no currency and
    no inflation profile is referenced, so FX and the deliberately-blank
    inflation grid are never resolved and neither is seeded.
    """
    declared = re.findall(
        r"@\{ Key = '(\w+)';\s+Value = \[double\](\S+);\s+Endpoint = '(\w+)' \}",
        _probe_code())
    assert declared == [
        ("base_year", "2026", "PCCM_ApplyTimeline"),
        ("project_start_year", "2027", "PCCM_ApplyTimeline"),
        ("duration_years", "3", "PCCM_ApplyTimeline"),
        ("discount_rate", "0.05", "PCCM_Calculate"),
    ], declared
    # AND NOTHING BEYOND THEM. A driver field, an FX rate or an inflation rate
    # would be the probe manufacturing business data.
    code = _probe_code()
    for banned in ("fx_rates", "inflation", "reporting_currency", "project_name",
                   "selected_confidence_level"):
        assert banned not in code, f"the probe seeds {banned}, which no command it runs needs"


def test_20a_the_discount_rate_is_supplied_as_a_real_number() -> None:
    """NOT HARD-CODED AROUND THE VALIDATION. modCalcResolve.IsRealNumber tests
    the VarType and refuses a numeric-looking STRING on purpose, so the value has
    to arrive as a genuine Double through the accepted setter."""
    code = _probe_code()
    assert "@{ Key = 'discount_rate';      Value = [double]0.05" in code
    setter = _function("Set-ProbeDeclaredInputs", code)
    assert "-Value ([double]$entry.Value)" in setter
    assert "Set-NamedValue -Workbook $Workbook" in setter
    # AND THE VALUE IS THE ACCEPTED FIXTURE'S, not a new one.
    timing = _src("timing", TIMING_SCENARIOS)
    assert "discount_rate = 0.05" in timing, (
        "the accepted fixture no longer declares the rate this reuses")


def test_20b_calculate_runs_where_its_prerequisites_are_satisfiable() -> None:
    """THE ORDERING IS THE FIXTURE. AddDriver writes a permanent ID, and
    ReadRegister reads every row whose id column is non-blank - so after an Add
    the register holds an identified driver with every other field empty and
    Calculate refuses on it. An empty driver set is valid by the contract's own
    wording, so Calculate runs BEFORE the Add commands and needs no invented
    business data."""
    code = _probe_code()
    # THE INVOCATIONS, not the stage labels. `-Endpoint 'PCCM_Calculate'` also
    # appears on Set-ProbeStage lines, so anchoring on it alone would find a
    # Calculate that is only being NAMED - and a mutation that deleted the call
    # outright would still satisfy the ordering.
    for endpoint in ("PCCM_ApplyTimeline", "PCCM_Calculate", "PCCM_AddCostLine",
                     "PCCM_AddRisk"):
        assert code.count(f"-Endpoint '{endpoint}' -Resolution $resolution") == 1, endpoint
    calculate = code.index("-Endpoint 'PCCM_Calculate' -Resolution $resolution")
    add_cost = code.index("-Endpoint 'PCCM_AddCostLine' -Resolution $resolution")
    add_risk = code.index("-Endpoint 'PCCM_AddRisk' -Resolution $resolution")
    timeline = code.index("-Endpoint 'PCCM_ApplyTimeline' -Resolution")
    assert timeline < calculate < add_cost < add_risk, (
        "Calculate no longer runs between ApplyTimeline and the Add commands")
    assert "An empty driver set is valid" in _probe() or \
        "AN EMPTY DRIVER SET IS VALID" in _probe()
    assert "AddDriver WRITES A PERMANENT ID" in _probe()


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
    setter = _function("Set-ProbeDeclaredInputs", _probe_code())
    assert "-Value ([double]$entry.Value)" in setter
    assert "Set-NamedValue -Workbook $Workbook" in setter


# ===========================================================================
# C2. CALCULATE IS JUDGED BY THE TABLES, NOT BY THE ANNOUNCEMENT
# ===========================================================================
def test_25_the_calc_tables_are_watched_around_calculate() -> None:
    """WINDOWS RUN 6 EXPOSED THIS PROBE'S OWN WEAK SPOT. It said the _Calc tables
    "are not watched above, so judge this one by its announcement" - while the
    probe's entire argument is that an announcement of success is not proof that
    a structural operation happened."""
    code = _probe_code()
    assert "function Get-ProbeCalcTables" in code
    assert "function Get-ProbeCalcShapes" in code
    assert "$calcBefore = @(Get-ProbeCalcShapes" in code
    assert "$calcAfter = @(Get-ProbeCalcShapes" in code
    before_at = code.index("$calcBefore = @(Get-ProbeCalcShapes")
    run_at = code.index("-Endpoint 'PCCM_Calculate' -Resolution $resolution")
    after_at = code.index("$calcAfter = @(Get-ProbeCalcShapes")
    assert before_at < run_at < after_at, "the shapes are not read either side of the call"
    assert "judge this one by its announcement" not in code, "the old wording is back"


def test_26_the_calc_table_authority_is_the_inspection() -> None:
    """NO LITERALS. The sheet, every table name and every row rule come from the
    same Gate-B inspection the rest of the probe reads."""
    code = _probe_code()
    calc = _inspection()["calc"]
    assert f"'{calc['sheet']}'" not in code, "the _Calc sheet name is hard-coded"
    for spec in calc["tables"].values():
        assert f"'{spec['table_name']}'" not in code, f"{spec['table_name']} is hard-coded"
    body = _function("Get-ProbeCalcTables", code)
    assert "-Name 'calc'" in body and "-Name 'sheet'" in body
    assert "-Name 'table_name'" in body and "-Name 'row_rule'" in body


def test_27_the_proof_is_predictive_not_merely_a_difference() -> None:
    """"SOMETHING CHANGED" IS NOT EVIDENCE OF THE CONTRACTED WORK. calc_years and
    calc_annual carry the row rule "one row per applied project year", so a
    Calculate that really ran ResizeBody leaves them holding exactly the applied
    duration - which only ResizeBody can produce."""
    code = _probe_code()
    body = _function("Get-ProbePerYearCalcTables", code)
    assert "[string]$entry.RowRule -eq 'one row per applied project year'" in body
    # THE RULE IS REALLY THE INSPECTION'S, and really names two tables.
    per_year = [name for name, spec in _inspection()["calc"]["tables"].items()
                if spec["row_rule"] == "one row per applied project year"]
    assert sorted(per_year) == ["calc_annual", "calc_years"], per_year
    assert "-ne $duration" in code
    assert "one per applied project year" in code


def test_28_a_refused_calculate_is_required_to_change_nothing() -> None:
    """AND IS NOT FAILED FOR IT. Demanding a resize from a command that refused
    would turn an honest refusal into a manufactured failure - which is the same
    mistake as calling any refusal a protection block."""
    code = _probe_code()
    assert "if ([string]$calculate.Outcome -eq 'SUCCEEDED') {" in code
    assert "A refusal is entitled to leave the tables alone." in code
    proof = code[code.index("$calcStructuralProof = 'NOT ESTABLISHED'"):]
    proof = proof[: proof.index("$addCost = Invoke-ProbeEndpoint")]
    for state in ("'OBSERVED'", "'CONTRADICTED'", "'NOT ESTABLISHED'"):
        assert state in proof, state


def test_29_an_announcement_the_shapes_contradict_cannot_reach_fine() -> None:
    """REQUIRED: no false PASS from the announcement alone."""
    code = _probe_code()
    assert "} elseif ($calcStructuralProof -eq 'CONTRADICTED') {" in code
    branch = code[code.index("} elseif ($calcStructuralProof -eq 'CONTRADICTED') {"):]
    branch = branch[: branch.index("} else {")]
    assert "$verdict = " not in branch, "the contradicted branch sets a verdict"
    assert "$verdictReason = " in branch
    assert branch.index("did not take") < len(branch)
    # AND IT IS WEIGHED BEFORE FINE.
    assert code.index("$calcStructuralProof -eq 'CONTRADICTED'") < \
        code.index("$verdict = 'PRODUCTION IS FINE UNDER PROTECTION'")


def test_2a_blocked_semantics_are_untouched() -> None:
    """REQUIRED: no weakening of BLOCKED. It still needs Excel's own sentence
    from an endpoint that was actually invoked."""
    code = _probe_code()
    assert ("$protectionBlocked = @(@($outcomes) | Where-Object "
            "{ Test-ProbeProtectionBlocked -Outcome $_ })") in code
    body = _function("Test-ProbeProtectionBlocked", code)
    assert "if (-not [bool]$Outcome.Invoked) { return $false }" in body
    assert "table features aren't available because the sheet is protected" in body.lower()
    gate = [line.strip() for line in code.splitlines()
            if "$protectionBlocked).Count -gt 0" in line]
    assert gate == ["if (@($protectionBlocked).Count -gt 0) {"], gate


def test_2b_the_add_commands_do_not_overclaim() -> None:
    """RUN 6: both SUCCEEDED with protection intact and NO watched shape change,
    which is correct for a fresh workbook - Stage A reserves 25 register rows so
    ListRows.Add never fires. Endpoint functionality under protection WAS
    observed; capacity expansion was NOT, and the probe says so."""
    source = _probe()
    assert "settles endpoint functionality under protection, NOT capacity expansion" in source
    assert "25 reserved rows" in source
    assert "does not settle" in source
    assert "no shape change on a fresh workbook is expected" in source


def test_2c_the_probe_never_releases_workbook_structure_protection() -> None:
    """RUN 6 SETTLED THE OPEN QUESTION: ApplyTimeline SUCCEEDED with structure
    protection True throughout, so ListColumns.Add on these tables does not need
    it released. The probe must not start releasing it either."""
    code = _probe_code()
    for banned in ("ThisWorkbook.Unprotect", ".Unprotect(", ".Unprotect ",
                   "ProtectStructure = ", "ProtectionRelease"):
        assert banned not in code, f"the probe releases protection: {banned}"
    # AND IT STILL READS THE STRUCTURE FLAG AS EVIDENCE.
    assert "Structure   = [bool]$Workbook.ProtectStructure" in code


# ===========================================================================
# D. READBACK BEFORE INVOCATION
# ===========================================================================
def test_30_every_input_is_read_back_before_the_endpoint() -> None:
    """REQUIRED CONTROL 6."""
    check = _function("Test-ProbeDeclaredInputs", _probe_code())
    assert "foreach ($entry in @($Inputs))" in check
    assert "Get-ProbeNamedValue -Workbook $Workbook -DefinedName $name" in check


def test_31_the_readback_checks_the_type_and_checks_it_first() -> None:
    """REQUIRED CONTROL 2. A value-only check passes a stringified number."""
    check = _function("Test-ProbeDeclaredInputs", _probe_code())
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
    problems_at = code.index("$inputProblems = @(Test-ProbeDeclaredInputs")
    throw_at = code.index("was invoked: ", problems_at)
    invoke_at = code.index("$timeline = Invoke-ProbeEndpoint", problems_at)
    assert problems_at < throw_at < invoke_at, "the readback does not gate the endpoint"
    # AND IT GATES EVERY ENDPOINT, not only the first: the throw sits above all
    # four Invoke-ProbeEndpoint calls, so a failed precondition means NONE ran.
    for endpoint in ("PCCM_ApplyTimeline' -Resolution", "PCCM_Calculate' -Resolution",
                     "PCCM_AddCostLine' -Resolution", "PCCM_AddRisk' -Resolution"):
        assert throw_at < code.index(endpoint, problems_at), endpoint
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
    # P10-RP. Production moved in the LATER runtime-protection reconciliation,
    # which is a different batch with its own authorisation. What this control
    # claims - that the Run-4 probe correction changed no production - still
    # holds, and is now proved by reversal rather than by an empty diff.
    import sys as _sys
    _sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_structural_window import (DECLARED_STRUCTURAL_WINDOW_CHANGES,
                                       strip_structural_window)
    declared = {f"pccm/src/vba/{name}" for name in DECLARED_STRUCTURAL_WINDOW_CHANGES}
    changed = [line for line in
               _git("diff", "--name-only", ACCEPTED, "--",
                    "pccm/src", "pccm/spec", "pccm/builder").splitlines() if line.strip()]
    for path in sorted(set(changed) & declared):
        name = Path(path).name
        current = strip_structural_window(
            name, (PCCM_ROOT / "src" / "vba" / name).read_bytes().decode("utf-8"))
        accepted = _git("show", f"{ACCEPTED}:{path}")
        assert current.replace("\r\n", "\n") == accepted.replace("\r\n", "\n"), path
    assert [p for p in changed if p not in declared] == [], changed


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
