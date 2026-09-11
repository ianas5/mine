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
        ("duration_years", "(Get-ProbeGrowthYears)", "PCCM_ApplyTimeline"),
        ("discount_rate", "0.05", "PCCM_Calculate"),
        ("duration_years", "(Get-ProbeShrinkYears)", "PCCM_ApplyTimeline"),
    ], declared
    # THE TWO DURATIONS LIVE IN ONE PLACE, and they must actually differ - a
    # shrink round that re-applied the same duration would delete nothing.
    code = _probe_code()
    grow = int(re.search(r"function Get-ProbeGrowthYears \{ return (\d+) \}", code).group(1))
    shrink = int(re.search(r"function Get-ProbeShrinkYears \{ return (\d+) \}", code).group(1))
    assert grow > shrink >= 1, (grow, shrink)
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
    assert code.count("-Endpoint 'PCCM_ApplyTimeline' -Resolution $resolution") == 2
    for endpoint in ("PCCM_AddCostLine", "PCCM_AddRisk"):
        assert code.count(f"-Endpoint '{endpoint}' -Resolution $resolution") == 1, endpoint
    # BOTH ROUNDS, and both between the timeline applies and the Add commands.
    # Counting matters: with only the shrink round left, an ordering check still
    # finds a Calculate in the right place and a mutation that deleted the growth
    # round walked straight through.
    assert code.count("Invoke-ProbeCalculateRound -Excel $excel") == 2, (
        "Calculate does not run for both the growth and the shrink round")
    assert "-ExpectedRows $growYears -Label 'growth'" in code
    assert "-ExpectedRows $shrinkYears -Label 'shrink'" in code
    calculate = code.index("Invoke-ProbeCalculateRound -Excel $excel")
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
    # INSIDE THE ROUND HELPER, which both rounds share so neither can drift.
    body = _function("Invoke-ProbeCalculateRound", code)
    assert "$before = @(Get-ProbeCalcShapes" in body
    assert "$after = @(Get-ProbeCalcShapes" in body
    before_at = body.index("$before = @(Get-ProbeCalcShapes")
    run_at = body.index("-Endpoint 'PCCM_Calculate' -Resolution $Resolution")
    after_at = body.index("$after = @(Get-ProbeCalcShapes")
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
    assert "-ne $ExpectedRows" in code
    assert "one per applied project year" in code


def test_28_a_refused_calculate_is_required_to_change_nothing() -> None:
    """AND IS NOT FAILED FOR IT. Demanding a resize from a command that refused
    would turn an honest refusal into a manufactured failure - which is the same
    mistake as calling any refusal a protection block."""
    code = _probe_code()
    body = _function("Invoke-ProbeCalculateRound", code)
    assert "if ([string]$outcome.Outcome -eq 'SUCCEEDED') {" in body
    assert "a refusal is entitled to leave the tables alone" in _probe().lower()
    for state in ("'OBSERVED'", "'CONTRADICTED'", "'NOT ESTABLISHED'"):
        assert state in body, state
    # AND THE SAME RULE GATES BOTH DELETE CLASSES.
    assert "if ([string]$shrinkTimeline.Outcome -eq 'SUCCEEDED') {" in code
    assert "if ([string]$shrinkCalculate.Outcome -eq 'SUCCEEDED') {" in code


def test_29_an_announcement_the_shapes_contradict_cannot_reach_fine() -> None:
    """REQUIRED: no false PASS from the announcement alone."""
    code = _probe_code()
    assert "} elseif ($shrinkRound.Proof -ne 'OBSERVED') {" in code
    branch = code[code.index("} elseif ($shrinkRound.Proof -ne 'OBSERVED') {"):]
    branch = branch[: branch.index("} else {")]
    assert "$verdict = " not in branch, "the contradicted branch sets a verdict"
    assert "$verdictReason = " in branch
    # AND BOTH ROUNDS ARE WEIGHED BEFORE FINE.
    fine = code.index("$verdict = 'PRODUCTION IS FINE UNDER PROTECTION'")
    assert code.index("$growRound.Proof -ne 'OBSERVED'") < fine
    assert code.index("$shrinkRound.Proof -ne 'OBSERVED'") < fine


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
    # AND IT STILL READS THE STRUCTURE FLAG AS EVIDENCE - the real Excel
    # property, now through the accepted retry boundary rather than bare. It is
    # not a proxy, not a cached value and not an assumption.
    assert "-Target $Workbook -Member 'ProtectStructure'" in code
    assert "Structure   = [bool](Invoke-ComRetryRead" in code


# ===========================================================================
# C3. THE DELETE PATH - THE CALL BENCHMARK RUN 3 ACTUALLY DIED ON
# ===========================================================================
# WINDOWS RUN 7 PROVED THE ADD DIRECTION AND ONLY THE ADD DIRECTION.
# ApplyTimeline grew three grids, Calculate grew tblCalcYears 1x3 -> 3x3 and
# tblCalcAnnual 1x8 -> 3x8, and protection was restored after every endpoint. But
# Benchmark Run 3 died on a ListRow.Delete(), and no successful ADD says anything
# about a DELETE. These controls pin the shrink round that closes that gap.
def test_2d_the_shrink_round_exists_and_changes_only_the_duration() -> None:
    code = _probe_code()
    assert "function Get-ProbeShrinkInputs" in code
    assert "$shrinkInputs = @(Get-ProbeShrinkInputs -Inspection $inspection)" in code
    assert "Set-ProbeDeclaredInputs -Workbook $wb -Inputs $shrinkInputs" in code
    body = _function("Get-ProbeShrinkInputs", code)
    keys = re.findall(r"@\{ Key = '(\w+)';", body)
    assert keys == ["duration_years"], (
        "the shrink round changes something other than the applied duration")
    # THE NAME STILL COMES FROM THE INSPECTION, never a literal.
    assert "inpDurationYears" not in code


def test_2e_the_shrink_input_is_read_back_and_type_checked() -> None:
    """THE SAME GATE AS EVERY OTHER INPUT. A shrink duration that did not land is
    probe instrumentation failure, never a production result."""
    code = _probe_code()
    assert "$shrinkProblems = @(Test-ProbeDeclaredInputs -Workbook $wb -Inputs $shrinkInputs)" in code
    assert "if (@($shrinkProblems).Count -gt 0) {" in code
    assert "the delete path was NOT" in code
    check = _function("Test-ProbeDeclaredInputs", code)
    assert "$actual -isnot [double]" in check


def test_2f_the_second_apply_timeline_is_invoked() -> None:
    """REMOVING IT WOULD LEAVE THE DELETE PATH UNEXERCISED."""
    code = _probe_code()
    assert code.count("-Endpoint 'PCCM_ApplyTimeline' -Resolution $resolution") == 2
    assert "$shrinkTimeline = Invoke-ProbeEndpoint" in code
    first = code.index("$timeline = Invoke-ProbeEndpoint")
    second = code.index("$shrinkTimeline = Invoke-ProbeEndpoint")
    assert first < second
    # AND THE SHRINK INPUT IS WRITTEN BETWEEN THEM.
    assert first < code.index("$shrinkInputs = @(Get-ProbeShrinkInputs") < second


def test_2g_column_delete_needs_the_columns_to_have_gone_down() -> None:
    """DIRECTION, NOT DIFFERENCE. A grid that merely CHANGED shape is not
    evidence of ListColumns.Delete, and a growth reading must never be accepted
    as a delete."""
    code = _probe_code()
    body = _function("Get-ProbeColumnDirection", code)
    assert "if ([int]$a.Columns -gt [int]$b.Columns) { $grew += $key }" in body
    assert "elseif ([int]$a.Columns -lt [int]$b.Columns) { $shrank += $key }" in body
    # EVERY GRID THAT GREW MUST SHRINK - derived from what was observed, so a
    # partial shrink that happened to touch one table is refused.
    assert "foreach ($key in @($growDirection.Grew)) {" in code
    assert "if (@($shrinkDirection.Shrank) -notcontains $key) { $missedShrink += $key }" in code
    assert "if ((@($growDirection.Grew).Count -gt 0) -and (@($missedShrink).Count -eq 0)) {" in code


def test_2h_row_delete_needs_the_rows_to_have_gone_down() -> None:
    """ACCEPTING 3 -> 3 WOULD CALL AN UNCHANGED TABLE DELETE EVIDENCE. Landing on
    the right number is not the same as having lost rows to get there."""
    code = _probe_code()
    body = _function("Invoke-ProbeCalculateRound", code)
    assert "if ([int]$now.Rows -lt [int]$b.Rows) {" in body
    assert "$shrank += [string]$b.Table" in body
    assert "RowsDeleted  = @($shrank)" in body
    # AND EVERY PER-YEAR TABLE MUST BE IN THAT LIST, not just one of them.
    assert "foreach ($table in @($perYearTables)) {" in code
    assert "if (@($shrinkRound.RowsDeleted) -notcontains [string]$table) {" in code
    assert "(@($missedRowDelete).Count -eq 0)" in code
    # AND THE GATE RESTS ON THE SHAPE PROOF, NOT ON THE ANNOUNCEMENT. The
    # endpoint outcome only decides whether a proof is DEMANDED; what makes it
    # OBSERVED is the round's own contracted shape.
    assert "if (([string]$shrinkRound.Proof -eq 'OBSERVED') -and" in code, (
        "row-delete evidence is granted on something other than the shape proof")
    gate = code[code.index("$rowDeleteProof = 'NOT ESTABLISHED'"):]
    gate = gate[: gate.index("$columnAddProof")]
    assert "$shrinkCalculate.Outcome -eq 'SUCCEEDED'" in gate, (
        "the proof is demanded of a Calculate that did not succeed")
    assert gate.count("$shrinkCalculate.Outcome") == 1, (
        "the announcement is consulted more than once; it decides only whether "
        "a proof is required")


def test_2i_the_delete_proofs_gate_the_fine_verdict() -> None:
    """FINE REQUIRES BOTH DIRECTIONS. Calling Benchmark Run 3 a harness defect on
    ADD evidence alone is exactly the overreach this refuses."""
    code = _probe_code()
    gate = "} elseif (($columnDeleteProof -ne 'OBSERVED') -or ($rowDeleteProof -ne 'OBSERVED')) {"
    assert gate in code
    branch = code[code.index(gate):]
    branch = branch[: branch.index("} elseif", 1)]
    assert "$verdict = " not in branch, "the missing-delete branch sets a verdict"
    assert "$verdictReason = " in branch
    assert code.index(gate) < code.index("$verdict = 'PRODUCTION IS FINE UNDER PROTECTION'")
    assert "Benchmark Run 3 died on a ListRow.Delete()" in code


def test_2j_the_four_evidence_classes_are_named_and_reported() -> None:
    """THE SEMANTICS ARE PINNED, and the success text distinguishes them."""
    code = _probe_code()
    for label in ("LISTCOLUMN ADD", "LISTCOLUMN DELETE", "LISTROW GROWTH", "LISTROW DELETE"):
        assert label in code, label
    fine = code[code.index("$verdict = 'PRODUCTION IS FINE UNDER PROTECTION'"):]
    fine = fine[: fine.index("}")]
    for label in ("LISTCOLUMN ADD: OBSERVED", "LISTCOLUMN DELETE: OBSERVED",
                  "LISTROW GROWTH: OBSERVED", "LISTROW DELETE: OBSERVED"):
        assert label in fine, label
    # EACH CLASS IS A WORD WITH THREE STATES, never a boolean.
    for variable in ("$columnAddProof", "$columnDeleteProof", "$rowDeleteProof"):
        assert f"{variable} = 'NOT ESTABLISHED'" in code, variable
        assert f"{variable} = 'OBSERVED'" in code, variable
        assert f"{variable} = 'CONTRADICTED'" in code, variable


def test_2k_protection_is_required_before_and_after_every_endpoint() -> None:
    """AND WORKBOOK STRUCTURE IS CHECKED AS A BOOLEAN, not read in a sentence."""
    code = _probe_code()
    assert "([int]$_.ProtectedAfter -ne [int]$_.TotalSheets) -or" in code
    assert "([int]$_.ProtectedBefore -ne [int]$_.TotalSheets) -or" in code
    assert "(-not [bool]$_.StructureBefore) -or (-not [bool]$_.StructureAfter)" in code
    assert "StructureBefore   = ([bool](Get-ProbeProperty -InputObject $protectionBefore" in code
    assert "StructureAfter    = ([bool](Get-ProbeProperty -InputObject $protectionAfter" in code
    # AND A LOST-PROTECTION READING STILL BLOCKS FINE.
    assert code.index("$lostProtection") < \
        code.index("$verdict = 'PRODUCTION IS FINE UNDER PROTECTION'")


# ===========================================================================
# C4. THE PROTECTION READER - WHAT ENDED WINDOWS RUN 8
# ===========================================================================
# Run 8 died three statements after Workbooks.Open, before the locked-cell
# control and before every production endpoint:
#
#   System.Management.Automation.PropertyNotFoundException
#   The property 'ProtectContents' cannot be found on this object.
#     at  if ($sheet.ProtectContents) { ... }
#
# The function was BYTE-IDENTICAL to the one Run 7 executed successfully, so the
# delete-path work did not introduce it. It is a latent defect in the reader.
def test_2l_the_protection_reader_proves_its_object_is_a_worksheet() -> None:
    """A BARE DEREFERENCE TURNS "no type information" INTO A SENTENCE NAMING A
    PROPERTY, which is why one Windows run could not diagnose it."""
    code = _probe_code()
    assert "function Assert-ProbeWorksheet" in code
    reader = _function("Get-ProbeProtectionState", code)
    assert "Assert-ProbeWorksheet -Candidate $sheet -Where $Where -Index $index" in reader
    # NOTHING SITS BETWEEN THE COLLECTION AND THE LOOP VARIABLE. A projection
    # into the foreach head would hand the reader objects that are not
    # Worksheets at all - which is the shape Run 8's failure looked like - and
    # an assertion placed after it would simply refuse every item.
    heads = [line.strip() for line in reader.splitlines()
             if line.strip().startswith("foreach (")]
    assert heads == ["foreach ($sheet in @($sheets)) {"], heads
    assert "$sheets = (Invoke-ComRetryRead -Target $Workbook -Member 'Worksheets'" in reader
    # THE ASSERTION COMES BEFORE ANY PROPERTY IS READ OFF THE ITEM.
    assert reader.index("Assert-ProbeWorksheet") < reader.index("-Member 'Name'")
    assert reader.index("Assert-ProbeWorksheet") < reader.index("-Member 'ProtectContents'")


def test_2m_the_membership_test_is_not_a_dereference() -> None:
    """READING .PSObject IS ALWAYS SAFE; reading a member that is not there is
    what StrictMode turns terminating. The check cannot be the crash."""
    body = _function("Assert-ProbeWorksheet", _probe_code())
    assert "$Candidate.PSObject.Properties['ProtectContents']" in body
    assert "$Candidate.PSObject.Properties['Name']" in body
    assert "if ($hasProtect -and $hasName) { return }" in body
    # AND IT NEVER DEREFERENCES THE PROPERTY IT IS TESTING FOR.
    assert "$Candidate.ProtectContents" not in body


def test_2n_a_non_worksheet_fails_explicitly_with_its_type() -> None:
    """REQUIRED: fail explicitly with diagnostic evidence including its actual
    PowerShell/.NET/COM type and the stage in which it occurred."""
    body = _function("Assert-ProbeWorksheet", _probe_code())
    assert "throw (" in body
    for fact in ("$Candidate.GetType().FullName", "$Candidate.PSObject.TypeNames",
                 "Marshal]::IsComObject($Candidate)", "$Candidate.PSObject.Properties).Count"):
        assert fact in body, f"the diagnosis omits {fact}"
    for label in (".NET type: ", "PSTypeNames: ", "IsComObject: ", "properties visible: "):
        assert label in body, label
    # THE STAGE AND THE ITEM ARE NAMED.
    assert "$Where + ': item ' + [string]$Index" in body
    # AND THE MESSAGE EXPLAINS WHAT THE EXCEPTION ACTUALLY MEANS.
    assert "exposes NO properties" in body


def test_2o_the_reader_never_assumes_a_protection_state() -> None:
    """REQUIRED: no assuming protected=true, no defaulting a missing property,
    no counting names only, no text matching, no swallowed exception."""
    reader = _function("Get-ProbeProtectionState", _probe_code())
    assert "-Member 'ProtectContents'" in reader, "the real property is no longer read"
    for banned in ("$true }", "-eq 'Protected'", "-match", "-like",
                   "catch { }", "SilentlyContinue"):
        assert banned not in reader, f"the reader weakens the check with {banned}"
    # BOTH BRANCHES COME FROM THE PROPERTY, and neither is a default.
    assert "if ($isProtected) { $protectedNames += $sheetName }" in reader
    assert "else              { $unprotectedNames += $sheetName }" in reader
    # AND A WORKBOOK THAT ENUMERATED NOTHING IS STILL REFUSED.
    assert "the workbook enumerated no worksheets" in reader


def test_2p_nothing_in_the_reader_swallows_a_property_not_found() -> None:
    """PropertyNotFoundException MUST STILL BE FATAL. Catching it is how a
    protection state nobody read becomes a protection state somebody reported."""
    code = _probe_code()
    # THE NAME MAY BE EXPLAINED, NEVER CAUGHT OR TESTED FOR. The diagnosis says
    # what the exception means; what is banned is a handler or a comparison that
    # would let the probe carry on past one.
    for banned in ("catch [System.Management.Automation.PropertyNotFoundException]",
                   "-is [System.Management.Automation.PropertyNotFoundException]",
                   "PropertyNotFoundException'", 'PropertyNotFoundException"'):
        assert banned not in code, f"the probe handles the exception rather than preventing it: {banned}"
    reader = _function("Get-ProbeProtectionState", code)
    assert "catch" not in reader.replace("catch { $hasProtect", ""), (
        "the protection reader catches something")
    # THE DIAGNOSTIC BUILDER'S CATCHES RECORD WHY, they do not discard.
    body = _function("Assert-ProbeWorksheet", code)
    for empty in ("catch { }", "catch {}"):
        assert empty not in body, "a diagnostic read is swallowed"
    assert body.count("'unreadable: ' + (Format-Err $_)") == 4


def test_2q_the_reader_reads_through_the_accepted_retry_boundary() -> None:
    """THE SAME BOUNDARY STAGE-B VERIFICATION ALREADY USES. Excel had just run
    Workbook_Open across 14 sheets, and a refused first call is a condition this
    machine has already produced once."""
    reader = _function("Get-ProbeProtectionState", _probe_code())
    for member in ("Worksheets", "Name", "ProtectContents"):
        assert f"-Member '{member}'" in reader, member
    assert reader.count("Invoke-ComRetryRead") >= 3
    # AND THE HELPER IS THE ACCEPTED ONE, not a local reimplementation.
    lifecycle = (PCCM_ROOT / "bootstrap" / "windows" / "com_lifecycle.ps1").read_text(encoding="utf-8")
    assert "function Invoke-ComRetryRead" in lifecycle
    assert "function Invoke-ComRetryRead" not in _probe_code(), (
        "the probe reimplements the accepted retry helper")


def test_2r_a_failed_protection_read_cannot_reach_a_verdict() -> None:
    """RUN 8 GOT THIS RIGHT AND IT MUST STAY RIGHT: the probe reported
    INCONCLUSIVE, the control NOT ATTEMPTED, and no endpoint was claimed."""
    code = _probe_code()
    # The reader throws; the outer handler leaves the verdict where it started.
    assert "$verdict = 'INCONCLUSIVE'" in code
    assert "the probe itself failed in stage" in code
    handler = code[code.index("the probe itself failed in stage"):]
    handler = handler[: handler.index("Write-ProbeLine")] if "Write-ProbeLine" in handler else handler
    assert "$verdict = 'PRODUCTION" not in handler
    # AND THE CONTROL STAYS NOT ATTEMPTED.
    assert "$controlResult = 'NOT ATTEMPTED'" in code
    assert "$controlDetail = 'the probe did not reach the locked-cell control'" in code


# ===========================================================================
# C5. THE ACCEPTANCE PREDICATES - THE RUN 9 CONTRADICTION
# ===========================================================================
# Windows Run 9 proved all four structural classes and then printed, in the same
# report:
#
#   C. NOT MET   PCCM_ApplyTimeline SUCCEEDED, created year columns, ...
#   STRUCTURAL INITIALISATION: NOT PROVEN - C not met.
#   ...
#   PRODUCTION IS FINE UNDER PROTECTION
#
# The runtime evidence was sound; the acceptance predicate was not.
def test_2s_each_criterion_binds_to_its_own_recorded_round() -> None:
    """THE ROOT CAUSE, REFUSED BY SHAPE. Criterion C used to select its round by
    ENDPOINT NAME and require exactly one match. The shrink round added a second
    outcome carrying the same name, so the count became 2 and C went false on a
    run whose growth apply had succeeded. A selector that is not unique is not a
    selector."""
    code = _probe_code()
    verdict = code[code.index("$criteria = New-Object"):]
    verdict = verdict[: verdict.index("$unmet = ")]
    # NO NAME-FILTERING OF THE OUTCOME LIST INSIDE THE CRITERIA.
    for banned in ("$_.Endpoint -eq", "$outcomes | Where-Object", "$timelineOutcome",
                   "$driverOutcomes"):
        assert banned not in verdict, (
            f"a criterion selects its round by filtering rather than by identity: {banned}")
    # EACH ONE NAMES THE ROUND IT IS ABOUT.
    assert "$timeline.Outcome" in verdict and "$timeline.StructuralEffect" in verdict
    assert "$addCost.Outcome" in verdict and "$addRisk.Outcome" in verdict
    # AND THE SHRINK ROUND IS NOT FOLDED INTO ANY OF THEM.
    assert "$shrinkTimeline" not in verdict, (
        "the shrink round is being read as part of an initialisation criterion")
    assert "$shrinkRound" not in verdict


def test_2t_criterion_c_is_the_first_growth_apply_and_says_so() -> None:
    code = _probe_code()
    assert "Key = 'C'; Text = 'the FIRST (growth) PCCM_ApplyTimeline SUCCEEDED" in code
    # AND IT CHECKS PROTECTION BOTH SIDES, INCLUDING STRUCTURE.
    verdict = code[code.index("Key = 'C';"):]
    verdict = verdict[: verdict.index("Key = 'D';")]
    for fact in ("$timeline.Outcome -eq 'SUCCEEDED'", "$timeline.StructuralEffect",
                 "$timeline.ProtectedBefore", "$timeline.ProtectedAfter",
                 "$timeline.StructureBefore", "$timeline.StructureAfter"):
        assert fact in verdict, f"criterion C no longer checks {fact}"


def test_2u_the_shrink_round_keeps_its_own_separate_representation() -> None:
    """LISTCOLUMN DELETE IS WHERE THE SHRINK EVIDENCE LIVES, and neither side is
    weakened by the other."""
    code = _probe_code()
    assert "$columnDeleteProof" in code and "$rowDeleteProof" in code
    assert "$shrinkTimeline.ShapesBefore" in code and "$shrinkTimeline.ShapesAfter" in code
    assert "LISTCOLUMN DELETE" in code


def test_2v_the_criteria_are_computed_before_the_verdict_uses_them() -> None:
    code = _probe_code()
    built = code.index("$criteria = New-Object")
    decided = code.index("if (@($protectionBlocked).Count -gt 0) {")
    assert built < decided, "the criteria are computed after the verdict is decided"
    assert "$structuralInitialisation = 'NOT PROVEN'" in code
    assert "if (@($unmet).Count -eq 0) { $structuralInitialisation = 'PROVEN' }" in code


def test_2w_fine_is_impossible_unless_every_criterion_is_met() -> None:
    """REQUIRED: a run must never print NOT PROVEN and FINE together."""
    code = _probe_code()
    gate = "} elseif ($structuralInitialisation -ne 'PROVEN') {"
    assert gate in code
    fine = code.index("$verdict = 'PRODUCTION IS FINE UNDER PROTECTION'")
    assert code.index(gate) < fine, "FINE is reachable without the criteria"
    branch = code[code.index(gate):]
    branch = branch[: branch.index("} else {")]
    assert "$verdict = " not in branch, "the not-proven branch sets a verdict"
    assert "not a self-consistent" in branch
    # AND THE FINE TEXT RECORDS THE CRITERIA TOO, so the artifact is readable.
    assert "STRUCTURAL INITIALISATION: PROVEN. " in code


def test_2x_fine_still_requires_all_four_evidence_classes() -> None:
    """THE CRITERIA ADD A REQUIREMENT; THEY DO NOT SUBSTITUTE FOR ONE."""
    code = _probe_code()
    fine = code.index("$verdict = 'PRODUCTION IS FINE UNDER PROTECTION'")
    for gate in ("} elseif (($columnDeleteProof -ne 'OBSERVED') -or ($rowDeleteProof -ne 'OBSERVED')) {",
                 "} elseif ($growRound.Proof -ne 'OBSERVED') {",
                 "} elseif ($shrinkRound.Proof -ne 'OBSERVED') {",
                 "} elseif ($structuralInitialisation -ne 'PROVEN') {"):
        assert gate in code, gate
        assert code.index(gate) < fine, gate
    # AND PROTECTION BOTH SIDES OF EVERY ENDPOINT STILL GATES IT.
    assert code.index("$lostProtection = @(") < fine


def test_2y_the_reporting_block_prints_what_the_verdict_used() -> None:
    """ONE COMPUTATION, ONE REPORT. The contradiction was possible because the
    criteria the report printed were computed where the verdict could not see
    them; recomputing them for printing would reintroduce exactly that."""
    code = _probe_code()
    report = code[code.index("-Action 'reporting the structural-initialisation criteria'"):]
    report = report[: report.index("STRUCTURAL EVIDENCE BY CLASS")]
    assert "$criteria = New-Object" not in report, "the report recomputes the criteria"
    assert "foreach ($criterion in @($criteria)) {" in report
    assert "if ($structuralInitialisation -eq 'PROVEN') {" in report
    assert code.count("$criteria = New-Object") == 1


def test_2z_the_predicate_correction_touched_nothing_that_drives_excel() -> None:
    """WHY RUN 9's EVIDENCE STANDS WITHOUT A RERUN.

    Every line this round changed sits AFTER the last endpoint invocation, in the
    verdict and reporting region. The sequence of Excel calls - which endpoints
    run, in what order, with what inputs, and which shapes and protection
    snapshots are read around them - is byte-identical to the revision Windows
    executed. So Run 9's runtime evidence describes exactly the code that is here
    now, and a rerun would produce a self-consistent REPORT rather than new
    evidence.
    """
    import subprocess as _sp

    executed = "04fcf82"
    diff = _sp.run(["git", "diff", executed, "--",
                    "pccm/bootstrap/windows/phase10_protection_probe.ps1"],
                   cwd=PCCM_ROOT.parent, check=True, stdout=_sp.PIPE, text=True).stdout
    hunks = [int(m.group(1)) for m in re.finditer(r"^@@ -\d+(?:,\d+)? \+(\d+)", diff, re.M)]
    assert hunks, "nothing changed since the executed revision, so this control is stale"

    lines = _probe().splitlines()
    last_excel = max(i for i, line in enumerate(lines, 1)
                     if ("Invoke-ProbeEndpoint" in line
                         or "Invoke-ProbeCalculateRound -Excel" in line))
    assert min(hunks) > last_excel, (
        f"a line at or before the last Excel interaction (line {last_excel}) changed: "
        f"earliest changed line is {min(hunks)}")
    # AND NOTHING THAT TOUCHES EXCEL APPEARS IN THE DIFF AT ALL.
    changed = [line[1:] for line in diff.splitlines()
               if line[:1] in "+-" and not line.startswith(("+++", "---"))]
    executable = [line for line in changed
                  if line.strip() and not line.strip().startswith("#")]
    for line in executable:
        for driver in ("$excel.", "$wb.", "Invoke-ProbeEndpoint", "Set-NamedValue",
                       "Get-ProbeAllShapes", "Get-ProbeCalcShapes",
                       "Get-ProbeProtectionState", ".Run("):
            assert driver not in line, f"the correction changes an Excel-driving line: {line.strip()}"


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
    # AND IT GATES EVERY ENDPOINT, not only the first: the throw sits above every
    # invocation, so a failed precondition means NONE ran.
    for marker in ("PCCM_ApplyTimeline' -Resolution", "Invoke-ProbeCalculateRound",
                   "PCCM_AddCostLine' -Resolution", "PCCM_AddRisk' -Resolution"):
        assert throw_at < code.index(marker, problems_at), marker
    # THE SHRINK ROUND HAS ITS OWN GATE, and it is not allowed to be softer: a
    # shrink input that did not land means the delete path was not exercised.
    shrink_at = code.index("$shrinkProblems = @(Test-ProbeDeclaredInputs")
    shrink_throw = code.index("the delete path was NOT", shrink_at)
    shrink_invoke = code.index("$shrinkTimeline = Invoke-ProbeEndpoint", shrink_at)
    assert shrink_at < shrink_throw < shrink_invoke, (
        "the shrink readback does not gate the shrink endpoint")
    assert "if (@($shrinkProblems).Count -gt 0) {" in code
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
    # FOUR NOW: the growth round's two precondition boundaries and the shrink
    # round's two. Every one of them says production has not run.
    assert code.count("production NOT invoked") == 4
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
