#!/usr/bin/env python3
"""Static validation of the Phase-7 sensitivity PERFORMANCE MEASUREMENT harness.

WHAT THIS FILE IS FOR
---------------------
`pccm/bootstrap/windows/phase7_timing_scenarios.ps1` cannot be executed on
Linux, and the task that produced it forbids running it on Windows. Everything
that CAN be proved without Excel is proved here, before anyone is asked to run
it:

  * it invokes the REAL public endpoint and not a surrogate;
  * the clock surrounds exactly that one call and nothing else;
  * the scenarios it claims to run are the scenarios it declares;
  * the addresses it reads are the contract's, not a second declaration;
  * the model it builds has genuinely varying drivers;
  * it modifies no production VBA and adds no timing code to any;
  * it leaves the accepted Gate-B harness alone, and Gate B does not run it;
  * it cannot orphan an Excel process.

WHAT IT DELIBERATELY DOES NOT ASSERT
------------------------------------
It does NOT freeze `pccm/src/vba` against the P7-4 head commit. That would pin
a MOMENT rather than a PROPERTY: the next production step would have to delete
the control to proceed, which is how a control becomes a formality. The durable
property is the one asserted below - the timing harness contains no VBA, writes
no `.bas`, and reaches into `pccm/src` for nothing but a hash - and the
byte-identity of production VBA at the moment of this measurement is reported as
evidence rather than frozen as a rule.
"""

from __future__ import annotations

import re
from pathlib import Path

import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent

BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
TIMING = BOOTSTRAP / "phase7_timing_scenarios.ps1"
DRIVER = BOOTSTRAP / "phase4_functional_test.ps1"
PHASE5 = BOOTSTRAP / "phase5_gate_b_scenarios.ps1"
PHASE6 = BOOTSTRAP / "phase6_gate_b_scenarios.ps1"
LIFECYCLE = BOOTSTRAP / "com_lifecycle.ps1"

SRC_VBA = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"

ENDPOINT = "PCCM_RunSensitivity"


def _text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _lines(path: Path) -> list[str]:
    return _text(path).replace("\r\n", "\n").split("\n")


def _executable(path: Path) -> str:
    """PowerShell source with the block comment and every `#` line removed.

    A rule about what the harness may not DO has to be checked against what it
    EXECUTES. Prose that explains why it never drives a Gate-B scenario set
    would otherwise trip the rule that it never drives one.
    """
    body = re.sub(r"<#.*?#>", "", _text(path), flags=re.S)
    return "\n".join(
        line for line in body.splitlines() if not line.lstrip().startswith("#")
    )


def _sim_contract() -> dict:
    with (SPEC / "sim_contract.yaml").open(encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def _workbook_spec() -> dict:
    with (SPEC / "workbook.yaml").open(encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def _column_number(letter: str) -> int:
    """A column letter as its 1-based ordinal.

    Spelled out because the sensitivity block now lives past column Z, where
    `ord(letter)` is not a column number and silently answers for the first
    character alone.
    """
    value = 0
    for char in str(letter).strip().upper():
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def _column_letter(number: int) -> str:
    out = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        out = chr(ord("A") + remainder) + out
    return out


def _function_body(source: str, name: str) -> str:
    """The text of one PowerShell function, brace-matched from its header.

    Brace matching, not a line count and not a "read until the next `function`":
    a helper that grew a nested scriptblock would silently fall out of a
    line-counted window, and the checks below would then be examining less than
    they claim to.
    """
    marker = "function " + name + " {"
    start = source.find(marker)
    assert start >= 0, f"{name} is not defined"
    depth = 0
    index = source.index("{", start)
    for position in range(index, len(source)):
        char = source[position]
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return source[start : position + 1]
    raise AssertionError(f"{name} is not brace-balanced")


# ===========================================================================
# 1. THE FILE, AND ITS PLACE IN THE HARNESS ESTATE
# ===========================================================================
def test_01_the_timing_harness_exists_and_is_a_windows_script() -> None:
    assert TIMING.is_file(), "the Phase-7 timing harness is missing"
    raw = TIMING.read_bytes()
    assert raw, "the timing harness is empty"
    assert b"\r\n" in raw, (
        "PowerShell scripts in this project are CRLF (.gitattributes declares "
        "`*.ps1 text eol=crlf`); the timing harness must match"
    )
    assert not re.search(rb"\r(?!\n)", raw), "a bare CR is not an accepted line ending"


def test_02_gate_b_does_not_run_the_timing_harness() -> None:
    """The measurement must not be able to affect a Gate-B result.

    The driver dot-sources the Phase-5 and Phase-6 scenario files. If it ever
    dot-sourced this one, a timing run would execute inside the accepted matrix
    and its Excel time would land in the middle of the evidence session.
    """
    for path in (DRIVER, PHASE5, PHASE6, LIFECYCLE):
        assert TIMING.name not in _text(path), (
            f"{path.name} references {TIMING.name}; the timing harness is "
            "additive and nothing in the accepted Gate-B estate may run it"
        )


def test_03_the_timing_harness_never_drives_a_gate_b_scenario_set() -> None:
    source = _executable(TIMING)
    for forbidden in ("Invoke-Phase5GateBScenarios", "Invoke-Phase6GateBScenarios",
                      "Add-Phase5Result", "Add-Phase6Result", "Add-Result"):
        assert forbidden not in source, (
            f"the timing harness calls {forbidden}; a measurement must not "
            "produce, record or imply a Gate-B result"
        )


def test_04_it_reuses_the_accepted_files_rather_than_reimplementing_them() -> None:
    source = _text(TIMING)
    for required in ("com_lifecycle.ps1", "phase5_gate_b_scenarios.ps1",
                     "phase6_gate_b_scenarios.ps1", "build_stage_b.ps1"):
        assert f"'{required}'" in source, (
            f"the timing harness must reuse {required} rather than carry its own copy"
        )
    assert "Set-Phase5Fixture" in source, (
        "the model must be established through the ACCEPTED fixture helper, so "
        "every driver is added by production's own Add endpoints"
    )
    assert "Save-Phase5LockedFxSeed" in source, (
        "the accepted fixture requires the locked FX seed to be captured on the "
        "untouched workbook before the first mutation"
    )


# ===========================================================================
# 2. THE COPIED PRIMITIVES CANNOT DRIFT
# ===========================================================================
COPIED_HELPERS = (
    "Write-RowObject",
    "Get-NamedValue",
    "Set-NamedValue",
    "Get-TableColumnNames",
    "Get-TableBody",
    "Set-TableCell",
    "Get-TableRowCount",
    "Add-BlankTableRow",
    "Remove-TableRow",
    "Get-IdColumnValues",
)


def test_05_every_copied_primitive_is_identical_to_the_drivers() -> None:
    """The copies are pinned, so the two files cannot become two readers.

    `phase4_functional_test.ps1` defines these at top level and cannot be
    dot-sourced (that would run the whole 103-case matrix), so the timing
    harness carries copies. A copy that drifts is a second, subtly different
    implementation of the same COM discipline - exactly the duplication the
    lifecycle policy exists to prevent - and this is what stops it.
    """
    driver = _text(DRIVER)
    timing = _text(TIMING)
    for name in COPIED_HELPERS:
        expected = _function_body(driver, name)
        actual = _function_body(timing, name)
        assert actual.replace("\r\n", "\n") == expected.replace("\r\n", "\n"), (
            f"{name} in the timing harness is not byte-identical to the "
            f"accepted driver's; one of the two has drifted"
        )


def test_06_the_copies_are_the_only_workbook_primitives_it_defines() -> None:
    """Nothing else in this file reads or writes the workbook by its own route.

    A private reader invented here would be a second contract for what a cell
    means, and its answers would look exactly like the accepted reader's.
    """
    declared = set(re.findall(r"^function ([A-Za-z][A-Za-z0-9-]*) \{", _text(TIMING), re.M))
    phase7_owned = {name for name in declared if name.startswith(("Get-Phase7", "New-Phase7",
                                                                 "Write-Phase7", "Format-Phase7",
                                                                 "Compare-Phase7"))}
    unexpected = declared - set(COPIED_HELPERS) - phase7_owned
    assert not unexpected, (
        "the timing harness defines helpers that are neither pinned copies nor "
        f"Phase-7 owned: {sorted(unexpected)}"
    )


# ===========================================================================
# 3. THE MEASURED CALL
# ===========================================================================
def _timed_region() -> list[str]:
    """The statements strictly between StartNew() and Stop(), structurally.

    Found by the stopwatch variable, not by a comment banner: a banner can be
    moved without moving the clock, and then the check would be reading a label
    instead of the thing it names.
    """
    lines = _lines(TIMING)
    starts = [i for i, line in enumerate(lines)
              if "$sensitivityWatch = [System.Diagnostics.Stopwatch]::StartNew()" in line]
    stops = [i for i, line in enumerate(lines) if "$sensitivityWatch.Stop()" in line]
    assert len(starts) == 1, "the sensitivity stopwatch must be started exactly once"
    assert len(stops) == 1, "the sensitivity stopwatch must be stopped exactly once"
    assert starts[0] < stops[0], "the sensitivity stopwatch is stopped before it is started"
    return lines[starts[0] + 1 : stops[0]]


def test_07_the_clock_surrounds_exactly_one_statement() -> None:
    body = [line.strip() for line in _timed_region()]
    body = [line for line in body if line and not line.startswith("#")]
    assert len(body) == 1, (
        "exactly one statement may sit inside the measured interval; found "
        f"{len(body)}: {body}"
    )


def test_08_the_measured_statement_is_the_real_public_endpoint() -> None:
    statement = [line.strip() for line in _timed_region()
                 if line.strip() and not line.strip().startswith("#")][0]
    assert statement == f"$excel.Run('{ENDPOINT}') | Out-Null", (
        "the measured statement must be the single Application.Run of the real "
        f"public {ENDPOINT} endpoint, and nothing else; found: {statement!r}"
    )


def test_09_the_endpoint_is_invoked_exactly_once_in_the_whole_file() -> None:
    """A second invocation outside the clock would make the number ambiguous."""
    invocations = re.findall(r"\$excel\.Run\('([A-Za-z_]+)'", _text(TIMING))
    assert invocations.count(ENDPOINT) == 1, (
        f"{ENDPOINT} must be invoked exactly once, inside the measured interval; "
        f"found {invocations.count(ENDPOINT)} invocation(s)"
    )


def test_10_the_automation_envelope_is_outside_the_measured_interval() -> None:
    """Opening the envelope and reading the announcement are not the endpoint.

    Both are COM round trips into VBA. Inside the clock they would be counted as
    sensitivity time; the report would then overstate the endpoint by however
    long the announcement took to read.
    """
    body = "\n".join(_timed_region())
    for excluded in ("PCCM_AutomationBegin", "PCCM_AutomationResult",
                     "PCCM_SimulationStatus", "Get-Phase6State"):
        assert excluded not in body, (
            f"{excluded} is inside the measured interval; only the endpoint may be"
        )


def test_11_nothing_else_is_timed_as_if_it_were_sensitivity() -> None:
    """The simulation and the fixture are measured on their OWN clocks.

    They are reported, because "how long did the run take" is a fair question -
    but they are separate stopwatches with separate labels, so no reader can
    mistake one for the other.
    """
    source = _text(TIMING)
    for other in ("$simulationWatch", "$fixtureWatch", "$calculateWatch"):
        assert f"{other} = [System.Diagnostics.Stopwatch]::StartNew()" in source, (
            f"{other} must exist so its cost is reported separately from the endpoint"
        )
    body = "\n".join(_timed_region())
    for other in ("$simulationWatch", "$fixtureWatch", "$calculateWatch", "$runStopwatch"):
        assert other not in body, f"{other} must not appear inside the measured interval"


def test_12_the_identity_invariants_are_captured_before_and_after() -> None:
    lines = _lines(TIMING)
    start = next(i for i, line in enumerate(lines)
                 if "$sensitivityWatch = [System.Diagnostics.Stopwatch]::StartNew()" in line)
    stop = next(i for i, line in enumerate(lines) if "$sensitivityWatch.Stop()" in line)
    before = [i for i, line in enumerate(lines) if "$before = Get-Phase6State" in line]
    after = [i for i, line in enumerate(lines) if "$after = Get-Phase6State" in line]
    assert len(before) == 1 and len(after) == 1, (
        "the run identity must be captured exactly once on each side of the measurement"
    )
    assert before[0] < start, "the BEFORE capture must precede the measured call"
    assert after[0] > stop, "the AFTER capture must follow the measured call"


def test_13_the_three_named_invariants_are_each_compared() -> None:
    """run_id, the AUTO nonce and result_digest, by name and not by a sweep.

    The sweep over every captured field is also there, and it is the stronger
    check - but a sweep that silently stopped covering these three would still
    pass, so the three the authorisation names are compared explicitly too.
    """
    source = _text(TIMING)
    for field in ("run_id", "result_digest", "consumed_auto_nonce", "next_auto_nonce"):
        assert f"Compare-Phase7Invariant" in source and f"'{field}'" in source, (
            f"{field} must be compared across the measured call by name"
        )
    assert "pending_auto_nonce" in source, (
        "the pending AUTO nonce cell must be compared across the measured call"
    )


def test_14_the_derived_status_rows_are_excluded_by_their_projected_group() -> None:
    """And excluded by the PROJECTION, not by a hard-coded pair of names.

    `RequireCurrentRun` asks for the simulation status before anything else, so
    the derived rows legitimately move. They are separated out by asking the
    inspection what group a row is in - the same question `Get-SimColumnFor`
    asks - rather than by this file deciding which two rows are special.
    """
    source = _text(TIMING)
    assert "$simInspection.sim_data.run_identity.groups" in source, (
        "the derived rows must be identified from the projection's own groups map"
    )
    assert "$group -eq 'derived'" in source, (
        "the exclusion must be on the projected group, not on row names"
    )
    for hard_coded in ("$fieldKey -eq 'simulation_status'",
                       "$fieldKey -eq 'status_evaluated_at'"):
        assert hard_coded not in source, (
            "the derived rows must not be named individually; the group is the rule"
        )


# ===========================================================================
# 4. THE SCENARIOS ARE WHAT THEY CLAIM TO BE
# ===========================================================================
def _declared_scenarios() -> list[dict]:
    body = _function_body(_text(TIMING), "Get-Phase7TimingScenarios")
    found = re.findall(
        r"Id\s*=\s*'([ABC])';\s*Title\s*=\s*'([^']+)';\s*"
        r"DriverCount\s*=\s*(\d+);\s*Iterations\s*=\s*(\d+)",
        body,
    )
    return [{"id": i, "title": t, "drivers": int(d), "iterations": int(n)}
            for i, t, d, n in found]


def test_15_exactly_the_three_authorised_scenarios_are_declared() -> None:
    scenarios = _declared_scenarios()
    assert [s["id"] for s in scenarios] == ["A", "B", "C"], (
        "the harness must declare exactly scenarios A, B and C, in ascending size"
    )
    assert [s["drivers"] for s in scenarios] == [20, 100, 300], (
        "the authorised driver counts are 20, 100 and 300"
    )
    assert [s["iterations"] for s in scenarios] == [10000, 10000, 10000], (
        "every scenario in this FIRST measurement runs 10,000 iterations"
    )


def test_16_there_is_no_hundred_thousand_iteration_scenario() -> None:
    """Explicitly excluded from the first measurement, so it is excluded here.

    Checked as a NUMBER in the executable text, not as the string '100000': a
    scenario added as 100000 would not be caught by looking for a word.
    """
    source = _text(TIMING)
    executable = "\n".join(line.split("#")[0] for line in source.replace("\r\n", "\n").split("\n"))
    assert not re.search(r"\b100000\b", executable), (
        "no 100,000-iteration scenario belongs in this first measurement"
    )
    for scenario in _declared_scenarios():
        assert scenario["iterations"] <= 10000, (
            f"scenario {scenario['id']} declares {scenario['iterations']} iterations"
        )


def test_17_every_scenario_mixes_cost_lines_and_risks() -> None:
    """The split is computed by the SAME rule the harness uses, not restated."""
    body = _function_body(_text(TIMING), "Get-Phase7CostLineCount")
    match = re.search(r"Ceiling\(\[double\]\$DriverCount \* ([0-9.]+)\)", body)
    assert match, "the cost/risk split must be a single arithmetic rule"
    share = float(match.group(1))
    assert 0.0 < share < 1.0, "the split must leave room for both kinds of driver"
    import math
    for scenario in _declared_scenarios():
        costs = math.ceil(scenario["drivers"] * share)
        risks = scenario["drivers"] - costs
        assert costs >= 1 and risks >= 1, (
            f"scenario {scenario['id']} would produce {costs} Cost Lines and "
            f"{risks} Risks; the model must contain both kinds"
        )


def test_18_only_the_driver_count_varies_across_the_scenarios() -> None:
    """One model builder, one parameter.

    If the builder took a second parameter the three scenarios would no longer
    be the same model at three sizes, and a timing difference could be caused by
    a shape difference nobody wrote down.
    """
    header = re.search(r"function New-Phase7TimingModel \{\s*\r?\n\s*param\(([^)]*)\)",
                       _text(TIMING))
    assert header, "New-Phase7TimingModel must declare its parameters"
    parameters = [p.strip() for p in header.group(1).split(",") if p.strip()]
    assert parameters == ["[int]$DriverCount"], (
        f"the model builder must vary by driver count alone; it takes {parameters}"
    )


def test_19_every_generated_driver_is_genuinely_varying() -> None:
    """Min < Most Likely < Max for every driver, from the builder's own literals.

    A model whose drivers were degenerate would time the zero-variance refusal
    path instead of the analysis, and the measurement would be of the wrong
    thing while looking entirely healthy.
    """
    body = _function_body(_text(TIMING), "New-Phase7TimingDriver")
    base = re.search(r"\$base = ([0-9.]+) \+ \(([0-9.]+) \* \[double\]\$Index\)", body)
    assert base, "the driver's base value must be a stated function of the index"
    intercept, slope = float(base.group(1)), float(base.group(2))
    assert intercept > 0 and slope > 0, "every driver must have a positive, varying base"

    most_likely = re.search(r"most_likely\s+= \[double\]\(\$base \* ([0-9.]+)\)", body)
    maximum = re.search(r"max_value\s+= \[double\]\(\$base \* ([0-9.]+)\)", body)
    assert re.search(r"min_value\s+= \$base", body), "min_value must be the base"
    assert most_likely and maximum, "most_likely and max_value must be multiples of the base"
    assert 1.0 < float(most_likely.group(1)) < float(maximum.group(1)), (
        "the model must satisfy Min < Most Likely < Max for every driver, so no "
        "driver is zero-variance and the ranked table is not mostly refusals"
    )


def test_20_risk_probabilities_are_strictly_between_zero_and_one() -> None:
    """A Risk that always or never occurs carries no occurrence variance."""
    body = _function_body(_text(TIMING), "New-Phase7TimingDriver")
    match = re.search(r"\(\(\(\(\$Index - 1\) % (\d+)\) \+ 1\) / ([0-9.]+)\)", body)
    assert match, "the risk probability must be a stated function of the index"
    modulus, divisor = int(match.group(1)), float(match.group(2))
    values = {((index - 1) % modulus + 1) / divisor for index in range(1, 400)}
    assert values, "the probability rule produced no values"
    assert min(values) > 0.0 and max(values) < 1.0, (
        f"risk probabilities must lie strictly inside (0, 1); got {sorted(values)}"
    )


def test_21_the_profiling_weights_match_the_timeline_and_sum_to_one() -> None:
    source = _text(TIMING)
    weights = re.search(r"profile_weights\s+= @\(([^)]*)\)", source)
    assert weights, "the model must declare profiling weights"
    values = [float(v.strip()) for v in weights.group(1).split(",")]
    duration = re.search(r"duration = (\d+)", source)
    assert duration, "the model must declare a timeline duration"
    assert len(values) == int(duration.group(1)), (
        "there must be exactly one profiling weight per project year"
    )
    assert sum(values) == 1.0, (
        f"the weights must sum to exactly 1.0 in binary64; they sum to {sum(values)!r}"
    )


def test_22_the_inflation_years_are_the_generated_columns() -> None:
    """BaseYear + 1 onwards, which is what production generates.

    Getting this wrong does not fail loudly: the fixture would throw "no
    generated inflation column for calendar year ...", after the workbook had
    already been half built.
    """
    source = _text(TIMING)
    base = int(re.search(r"base_year = (\d{4})", source).group(1))
    duration = int(re.search(r"duration = (\d+)", source).group(1))
    start = int(re.search(r"start_year = (\d{4})", source).group(1))
    assert start == base + 1, (
        "the accepted shape puts the first project year at BaseYear + 1, where "
        "the first generated inflation column is"
    )
    expected = {str(base + 1 + offset) for offset in range(duration)}
    declared = set(re.findall(r"'(\d{4})' = [0-9.]+", source))
    assert declared == expected, (
        f"the inflation rates cover {sorted(declared)}; the generated columns are "
        f"{sorted(expected)}"
    )


# ===========================================================================
# 5. THE ADDRESSES ARE THE CONTRACT'S
# ===========================================================================
def _geometry_block(name: str) -> str:
    source = _text(TIMING)
    start = source.index(f"$script:{name} = [pscustomobject]@{{")
    depth = 0
    index = source.index("{", start)
    for position in range(index, len(source)):
        if source[position] == "{":
            depth += 1
        elif source[position] == "}":
            depth -= 1
            if depth == 0:
                return source[start : position + 1]
    raise AssertionError(f"{name} is not brace-balanced")


def test_23_the_sensitivity_block_geometry_is_the_contracts() -> None:
    block = _geometry_block("Phase7SensitivityGeometry")
    records = _sim_contract()["sim_data"]["sensitivity_records"]

    assert re.search(r"HeaderRow\s+= " + str(records["header_row"]) + r"\b", block)
    assert re.search(r"FirstRecordRow\s+= " + str(records["first_record_row"]) + r"\b", block)

    stamp = records["stamp"]
    for bank, column in stamp["bank_value_columns"].items():
        assert f"'{bank}' = '{column}'" in block, (
            f"the stamp column for bank {bank} must be {column}, as the contract says"
        )
    for field in stamp["fields"]:
        assert re.search(rf"{field['key']}\s+= {field['row']}\b", block), (
            f"the stamp row for {field['key']} must be {field['row']}"
        )

    status = [column for column in records["columns"] if column["key"] == "status"][0]
    banks = records["banks"]
    first = records["columns"][0]
    offset = _column_number(status["column"]) - _column_number(first["column"])
    for bank, span in banks.items():
        expected = _column_letter(_column_number(span["first_column"]) + offset)
        assert re.search(rf"'{bank}' = '{expected}'", block), (
            f"the status column for bank {bank} is {expected}: the block starts at "
            f"{span['first_column']} and status sits {offset} column(s) into it"
        )
        assert f"'{bank}' = '{span['first_column']}'" in block, (
            f"the id/stamp column for bank {bank} is the block's first column, "
            f"{span['first_column']}"
        )


def test_24_the_status_labels_are_productions_own() -> None:
    """The labels are Private constants in modSimPostReport, so they are read
    from there. A harness that spelled 'n/a - no variance' its own way would
    count every refusal as an unrecognised label and report zero of them."""
    block = _geometry_block("Phase7SensitivityGeometry")
    module = (SRC_VBA / "modSimPostReport.bas").read_text(encoding="utf-8")
    for constant, key in (("SENSITIVITY_NO_VARIANCE_LABEL", "NoVarianceLabel"),
                          ("SENSITIVITY_RANKED_LABEL", "RankedLabel")):
        match = re.search(
            rf'Private Const {constant} As String = "([^"]*)"', module)
        assert match, f"{constant} is not declared in modSimPostReport"
        assert re.search(rf"{key}\s+= '{re.escape(match.group(1))}'", block), (
            f"{key} must be production's {constant}, which is "
            f"'{match.group(1)}'"
        )
    published = re.search(r'const\("SIM_SENSITIVITY_PUBLISHED", "([^"]+)"\)',
                          (PCCM_ROOT / "builder" / "pccm_builder" / "sim_emit.py")
                          .read_text(encoding="utf-8"))
    assert published, "the published marker is not projected by sim_emit"
    assert re.search(rf"PublishedMarker\s+= '{re.escape(published.group(1))}'", block)


def test_25_the_sensitivity_sheet_geometry_is_the_workbook_specs() -> None:
    block = _geometry_block("Phase7SensitivitySheet")
    shell = _workbook_spec()["phase6_shell"]["sensitivity"]
    assert f"Sheet              = '{shell['sheet']}'" in block
    assert re.search(r"AvailabilityRow\s+= " + str(shell["availability_row"]) + r"\b", block)
    assert re.search(r"HeaderRow\s+= " + str(shell["header_row"]) + r"\b", block)
    assert re.search(r"FirstRow\s+= " + str(shell["first_row"]) + r"\b", block)
    assert re.search(r"RowWindow\s+= " + str(shell["row_window"]) + r"\b", block)
    assert f"AvailabilityColumn = '{shell['label_column']}'" in block
    assert f"FirstColumn        = '{shell['first_column']}'" in block
    last = shell["columns"][-1]["column"]
    assert f"LastColumn         = '{last}'" in block


def test_26_the_pinned_phase6_inspection_is_not_extended() -> None:
    """The addresses are declared locally FOR A REASON, and the reason holds.

    `phase6_gate_b_inspection.json` carries a pinned SHA-256 that the Phase-6
    fixture-integrity checks depend on. Projecting the Phase-7 block into it
    would move that hash and break the historical Phase-6 identity controls.
    """
    projection = (PCCM_ROOT / "builder" / "pccm_builder" / "sim_inspection.py").read_text(
        encoding="utf-8")
    assert "sensitivity_records" not in projection, (
        "the Phase-6 inspection projection must not grow a Phase-7 section; that "
        "would move a pinned artefact hash"
    )
    assert "phase6_gate_b_inspection.json" in _text(TIMING), (
        "the timing harness still reads the accepted Phase-6 projection for every "
        "address that projection already carries"
    )


# ===========================================================================
# 6. PRODUCTION IS UNTOUCHED, AND EXCEL CANNOT BE ORPHANED
# ===========================================================================
def test_27_the_timing_harness_contains_no_vba_and_writes_none() -> None:
    """No timing code is added to production, and none can be.

    The harness reaches into `pccm/src/vba` for exactly one thing - a hash - and
    has no writing verb anywhere near it.
    """
    source = _text(TIMING)
    for verb in ("Set-Content", "Out-File", "Add-Content", "CodeModule", "AddFromString",
                 "VBProject", "Import "):
        if verb == "Set-Content":
            # The report is the one thing written, and it is written to the
            # disposable temp directory.
            occurrences = re.findall(r"Set-Content -LiteralPath (\S+)", source)
            assert occurrences == ["$script:Phase7ReportPath"], (
                f"Set-Content may only write the report; it targets {occurrences}"
            )
            continue
        assert verb not in source, (
            f"the timing harness must not contain {verb!r}: it neither builds nor "
            "modifies a VBA project"
        )
    assert ".bas" in source, "the module identities are read from the .bas files"
    assert not re.search(r"(Remove-Item|Set-Content|Copy-Item)[^\r\n]*src[/\\]vba", source), (
        "nothing in the timing harness may write into pccm/src/vba"
    )


def test_28_the_workbook_is_never_saved() -> None:
    """So no measurement carries a save cost, and no build output is mutated."""
    source = _text(TIMING)
    assert "$wb.Close($false)" in source, "the workbook must be closed WITHOUT saving"
    for forbidden in ("$wb.Save(", ".SaveAs(", "$wb.Saved ="):
        assert forbidden not in source, (
            f"{forbidden} would either mutate the workbook or add a save cost to a "
            "measurement"
        )


def test_29_shutdown_is_the_accepted_lifecycle_and_cannot_orphan_excel() -> None:
    source = _text(TIMING)
    for required in ("New-ReleaseLedger", "Invoke-NamedRelease", "$excel.Quit()",
                     "Wait-ExcelExit", "Invoke-EmergencyExcelCleanup",
                     "Get-PreExistingExcelPids", "Get-ExcelIdentity"):
        assert required in source, (
            f"{required} is part of the accepted COM lifecycle and must be used"
        )
    assert "Stop-Process" not in source, (
        "the timing harness must never terminate a process itself; only "
        "Invoke-EmergencyExcelCleanup may, and only for an identity it verified"
    )
    assert "FinalReleaseComObject" not in source, (
        "FinalReleaseComObject is prohibited by the lifecycle policy"
    )
    # The shutdown must be unconditional. A `finally` is the only construct that
    # survives the measurement raising, and a raised measurement is exactly when
    # an orphan would be created.
    assert re.search(r"\}\s*finally\s*\{[^{]*New-ReleaseLedger", source, re.S), (
        "the shutdown must run from a finally block, so a failed measurement "
        "still closes Excel"
    )


def test_30_a_workbook_that_cannot_be_attributed_is_refused() -> None:
    """Fail closed, before Excel is started.

    "The timing workbook contains the current Phase-7 projection" is a claim
    about bytes. With a modified `pccm/src`, `pccm/spec` or `pccm/builder` there
    is no revision to attribute those bytes to, and a measurement nobody can
    attribute is not evidence.
    """
    source = _text(TIMING)
    body = _function_body(source, "Get-Phase7SourceRevision")
    for pathspec in ("'pccm/src'", "'pccm/spec'", "'pccm/builder'"):
        assert pathspec in body, f"{pathspec} must be checked for modification"
    assert "rev-parse HEAD" in body, "the source revision must be recorded"
    # THE GUARD IS READ AS A WHOLE LINE, not as a substring anywhere in it.
    # `if ($false -and $revision.Dirty.Count -gt 0)` contains the comparison and
    # tests nothing; a check that searched for the comparison would call that
    # refusal intact.
    guards = [line.strip() for line in _lines(TIMING)
              if "$revision.Dirty.Count" in line]
    assert guards == ["if ($revision.Dirty.Count -gt 0) {"], (
        f"the dirty-tree guard must be exactly that comparison; found {guards}"
    )
    refusal = source.index("$revision.Dirty.Count -gt 0")
    opening = source.index("New-Object -ComObject Excel.Application")
    assert refusal < opening, (
        "the refusal must happen BEFORE an Excel instance is created"
    )
    assert "exit 1" in source[refusal:opening], "the refusal must stop the run"


def test_31_the_module_identities_use_the_accepted_canonicalisation() -> None:
    """The same hash the accepted Phase-6 projection identity uses.

    A second canonicaliser would answer the same question differently - a bare
    CR admitted here and refused there - and the two identities would stop being
    comparable.
    """
    source = _text(TIMING)
    assert "Get-Phase6CanonicalModuleHash" in source, (
        "module identity must be taken with the accepted canonical hash"
    )
    assert "System.Security.Cryptography.SHA256" not in source, (
        "the timing harness must not carry its own hasher"
    )
    body = _function_body(source, "Get-Phase7ModuleIdentities")
    assert "$Manifest.vba.modules" in body, (
        "every module the workbook contains must be hashed, from the manifest's "
        "own list, so a module cannot be omitted by being forgotten"
    )
    assert "$module.generated" in body, (
        "a generated module comes from the disposable build copy and a "
        "hand-written one from the repository; the two origins are not the same "
        "claim and must not be conflated"
    )


# ===========================================================================
# 7. THE BOUND IS REAL, AND HONESTLY DESCRIBED
# ===========================================================================
def test_32_the_bound_is_an_entry_gate_and_says_so() -> None:
    source = _text(TIMING)
    assert "$SensitivityBudgetSeconds" in source and "$TotalBudgetSeconds" in source, (
        "both bounds must be parameters, so a reviewer can see and change them"
    )
    # THE GATE IS A CLOSED LOOP, and each half is checked where it lives.
    #
    #   ARMED    a measured time over the budget must ASSIGN $gateReason
    #   ENFORCED a set $gateReason must SKIP the scenario
    #
    # Either half alone is decoration: an armed gate nothing reads changes
    # nothing, and an enforced gate nothing arms never fires.
    armed = re.search(
        r"if \(\$sensitivityWatch\.Elapsed\.TotalSeconds -gt "
        r"\[double\]\$SensitivityBudgetSeconds\) \{\s*\r?\n\s*\$gateReason =",
        source)
    assert armed, (
        "a measured sensitivity time over the budget must set $gateReason; "
        "without that assignment the budget parameter is decoration"
    )
    enforced = re.search(
        r"if \(-not \[string\]::IsNullOrEmpty\(\$gateReason\)\) \{"
        r"(?:.(?!\n        \}))*?continue",
        source, re.S)
    assert enforced, (
        "a set $gateReason must skip the remaining scenarios; without the skip "
        "the gate records a reason and enters the scenario anyway"
    )
    assert "NOT ENTERED" in source, (
        "a scenario stopped by the gate must be reported as NOT ENTERED, not "
        "silently omitted"
    )
    assert "ENTRY-GATE, not interruption" in source, (
        "the harness must state plainly that it cannot interrupt a running "
        "synchronous Application.Run, rather than implying a hard timeout it "
        "does not have"
    )


def test_33_the_scenarios_run_smallest_first_and_report_as_they_go() -> None:
    """So the cheap evidence survives whatever happens to the expensive run."""
    scenarios = _declared_scenarios()
    assert [s["drivers"] for s in scenarios] == sorted(s["drivers"] for s in scenarios), (
        "scenarios must be declared in ascending size, so the gate protects the "
        "expensive ones and never the cheap one"
    )
    write_through = _function_body(_text(TIMING), "Write-Phase7Line")
    assert "Set-Content" in write_through, (
        "every reported line must be flushed to disk as it is produced; a run "
        "that is stopped must still leave its completed measurements behind"
    )


def test_34_nothing_is_capped_or_subsampled_to_improve_a_number() -> None:
    """Evidence first, architecture decision second.

    The authorisation is explicit that a high timing must not be answered by
    quietly introducing a cap. There is no Top-N and no subsampling anywhere in
    the measurement path, and the report says so.
    """
    executable = _executable(TIMING)
    for forbidden in ("Select-Object -First", "Select-Object -Last", "[Math]::Min(",
                      "TopN", "Subsample", "-Sample"):
        assert forbidden not in executable, (
            f"{forbidden} has no place in the executable path of a measurement "
            "harness: every driver the model declares must reach the endpoint, and "
            "every record the endpoint produced must be counted"
        )
    # The record sweep runs to the STAMP'S OWN record_count. A constant ceiling
    # here would under-report the very scenario the ceiling question is about.
    counts = _function_body(_text(TIMING), "Get-Phase7SensitivityStatusCounts")
    assert "$offset -lt $RecordCount" in counts, (
        "the status sweep must run to the record count production wrote, not to a "
        "ceiling the harness chose"
    )
    assert "nothing was subsampled to make a number look better" in _text(TIMING), (
        "the report must state that no cap or subsampling was applied"
    )


def test_35_the_report_captures_every_required_observation() -> None:
    """The authorisation's capture list, item by item.

    Checked as REPORT LABELS the file emits, because the deliverable is a
    report a person reads: an observation that is measured but never printed is
    not captured.
    """
    source = _text(TIMING)
    required = (
        "git HEAD",                     # source / build revision
        "SCENARIO ",                    # scenario identifier
        "drivers in book",              # driver count
        "Cost Lines in book",           # cost line count
        "Risks in book",                # risk count
        "iterations control",           # iteration count
        "PCCM_RunSimulation   : ",      # simulation success
        "status before",                # current status before sensitivity
        "sensitivity result",           # success or refusal
        "sensitivity records",          # record count
        "eligible ranked",              # eligible ranked-driver count
        "zero variance",                # zero-variance count
        "PCCM_RunSensitivity ELAPSED",  # the measurement itself
        "status after",                 # final simulation status
        "identity invariants",          # run_id / nonce / digest
        "Sensitivity sheet",            # the materialisation
        "simulation time",              # reported separately
    )
    missing = [label for label in required if label not in source]
    assert not missing, f"the report never emits: {missing}"


# ===========================================================================
# 8. THE SCENARIO LOOP REACHES EVERY SCENARIO IT DECLARES
# ===========================================================================
# Runtime run 1 of this harness measured scenario A and stopped. B and C were
# never entered and nothing said why: `-Scenario A` had FILTERED THEM OUT OF
# THE COLLECTION before the loop, so there was no scenario left to skip and no
# skip to report. The loop itself was never defective.
#
# The correction made selection a reported outcome rather than a filter, and
# these controls hold that shape. They are structural: the loop body is
# brace-matched and every control transfer in it is classified by the construct
# that encloses it, so a `continue` that belongs to the identity sweep is not
# mistaken for one that skips a scenario, and a new one that does skip a
# scenario cannot be added without being named here.
LOOP_HEADER = "foreach ($case in $declared) {"

CONSTRUCT = re.compile(
    r"^(foreach|while|do|for|if|elseif|else|try|catch|finally|switch)\b")
TRANSFER = re.compile(r"^(continue|break|return|exit|throw)\b")


def _scenario_loop() -> tuple[int, list[tuple[int, str, list[str]]]]:
    """(first body line, [(line no, text, enclosing-construct stack)]).

    The stack is what makes the classification honest. `continue` at depth 3
    inside `foreach ($fieldKey ...)` advances THAT loop; only a transfer whose
    enclosing stack contains no inner loop acts on the scenario loop.
    """
    lines = _lines(TIMING)
    starts = [i for i, line in enumerate(lines) if line.strip() == LOOP_HEADER]
    assert len(starts) == 1, (
        f"there must be exactly one scenario loop, written {LOOP_HEADER!r}; "
        f"found {len(starts)}"
    )
    start = starts[0]
    depth = 0
    end = None
    for index in range(start, len(lines)):
        depth += lines[index].count("{") - lines[index].count("}")
        if depth == 0 and index > start:
            end = index
            break
    assert end is not None, "the scenario loop is not brace-balanced"

    annotated: list[tuple[int, str, list[str]]] = []
    stack: list[str] = []
    for offset, line in enumerate(lines[start + 1 : end]):
        stripped = line.strip()
        annotated.append((start + 2 + offset, stripped, list(stack)))
        for _ in range(line.count("{")):
            head = stripped.split("{")[0].strip()
            match = CONSTRUCT.match(head)
            stack.append(match.group(1) if match else "block")
        for _ in range(line.count("}")):
            if stack:
                stack.pop()
    return start + 2, annotated


def _scenario_level(annotated) -> list[tuple[int, str]]:
    """Transfers that act on the SCENARIO loop, not on a loop nested in it."""
    loops = {"foreach", "while", "do", "for"}
    return [(number, text) for number, text, stack in annotated
            if TRANSFER.match(text) and not (loops & set(stack))]


def test_36_the_loop_iterates_the_declared_set_and_nothing_narrower() -> None:
    """The collection the loop walks must be the full declared set.

    This is the control that would have caught runtime run 1. `$selected`, a
    pre-filtered copy, is exactly the shape that let two scenarios vanish
    without a line of explanation.
    """
    source = _executable(TIMING)
    assert "$declared = @(Get-Phase7TimingScenarios)" in source, (
        "the declared set must come straight from the scenario table"
    )
    assignments = re.findall(r"^\s*\$declared\s*=", source, re.M)
    assert len(assignments) == 1, (
        f"$declared must be assigned exactly once and never re-filtered; found "
        f"{len(assignments)} assignment(s)"
    )
    assert not re.search(r"\$declared\s*=\s*@\(\$declared", source), (
        "the declared set must never be narrowed in place"
    )
    # And selection must not be able to shrink it by another route.
    assert "$selected " not in source and "$selected)" not in source, (
        "a pre-filtered scenario collection must not exist; selection is "
        "reported per scenario inside the loop"
    )


def test_37_every_way_out_of_the_scenario_loop_is_named() -> None:
    """Exactly three, all `continue`, all of which report themselves.

    Nothing may leave the loop silently, and nothing may end the RUN from
    inside it: a `break` would stop the remaining scenarios with no line, and a
    `return`/`exit` would end the script before the summary and the shutdown.
    """
    _, annotated = _scenario_loop()
    transfers = _scenario_level(annotated)
    assert [text for _, text in transfers] == ["continue", "continue", "continue"], (
        "the scenario loop must have exactly three scenario-level control "
        f"transfers, all `continue`; found {transfers}"
    )
    for number, text, _ in annotated:
        assert not re.match(r"^(break|return|exit)\b", text), (
            f"line {number} ({text!r}) can end the run from inside the scenario "
            "loop; the loop must always fall through to the summary and the "
            "shutdown"
        )
    # Each `continue` must be preceded by a line that reports the skip, so a
    # skipped scenario is always visible in the report.
    reported = 0
    for index, (number, text, _) in enumerate(annotated):
        if text != "continue":
            continue
        window = " ".join(entry[1] for entry in annotated[max(0, index - 12): index])
        if "NOT SELECTED" in window or "NOT ENTERED" in window or "ABANDONED" in window:
            reported += 1
    assert reported == 3, (
        f"every scenario-level `continue` must report itself; {reported} of 3 do"
    )


def test_38_all_three_scenarios_are_entered_when_each_is_successful() -> None:
    """The positive path, proved by exhausting the negative ones.

    The loop walks the declared set (test_36), the only ways out are the three
    named skips (test_37), and each of those three is guarded by a condition
    that a successful, selected, under-budget scenario does not satisfy:

        not selected     - false for every scenario when -Scenario is All
        gate armed       - $gateReason is empty until something arms it
        simulation failed- the announcement was OK|

    So with three successful under-budget scenarios in scope, the loop
    necessarily reaches all three.
    """
    _, annotated = _scenario_loop()
    guards = [text for _, text, stack in annotated
              if not stack and text.startswith("if (")]
    assert guards[0] == "if ($selectedIds -notcontains [string]$case.Id) {", (
        f"scope must be the loop's first decision; it is {guards[0]!r}"
    )
    assert guards[1] == "if ([string]::IsNullOrEmpty($gateReason)) {", (
        "the budget check must follow scope, so an unselected scenario neither "
        f"spends the budget nor arms the gate; found {guards[1]!r}"
    )
    assert guards[2] == "if (-not [string]::IsNullOrEmpty($gateReason)) {", (
        f"the gate must be enforced next; found {guards[2]!r}"
    )
    # Nothing may arm the gate on the success path. Counted INSIDE the loop, so
    # the one initialisation to '' outside it is not mistaken for an arming.
    source = _executable(TIMING)
    inside = [text for _, text, _ in annotated if text.startswith("$gateReason = ")]
    assert len(inside) == 3, (
        f"$gateReason must be armed in exactly three places inside the loop - the "
        f"total-budget check, the simulation-failure path and the over-budget "
        f"check; found {len(inside)}: {inside}"
    )
    outside = re.findall(r"^\$gateReason = ''$", source, re.M)
    assert len(outside) == 1, (
        "the gate must start disarmed, initialised exactly once before the loop"
    )
    # -Scenario All must select everything, so nothing is out of scope.
    assert "if ($Scenario -ne 'All') { $selectedIds = @([string]$Scenario) }" in source, (
        "-Scenario All must leave every declared scenario selected"
    )
    assert "$selectedIds = @($declared | ForEach-Object { [string]$_.Id })" in source, (
        "the default selection must be every declared scenario's id"
    )


def test_39_an_over_budget_scenario_stops_the_later_ones_and_says_so() -> None:
    """Armed after the measurement, enforced before the next scenario's work."""
    first, annotated = _scenario_loop()
    numbers = {text: number for number, text, _ in annotated}
    source = _text(TIMING)
    armed = re.search(
        r"if \(\$sensitivityWatch\.Elapsed\.TotalSeconds -gt "
        r"\[double\]\$SensitivityBudgetSeconds\) \{\s*\r?\n\s*\$gateReason =", source)
    assert armed, "an over-budget measured time must arm the gate"

    enforced_line = next(number for number, text, stack in annotated
                         if not stack and text == "if (-not [string]::IsNullOrEmpty($gateReason)) {")
    armed_line = next(number for number, text, _ in annotated
                      if text.startswith("$gateReason = ('scenario '"))
    assert enforced_line < armed_line, (
        "the gate must be enforced at the TOP of the next iteration, after "
        "being armed at the bottom of this one"
    )
    # The enforced branch reports NOT ENTERED and records an unentered row.
    block = [text for number, text, _ in annotated
             if enforced_line <= number <= enforced_line + 12]
    assert any("NOT ENTERED: " in text for text in block), (
        "a gated scenario must be reported as NOT ENTERED with its reason"
    )
    assert any("Selected = $true; Entered = $false" in text for text in block), (
        "a gated scenario must be recorded as selected but not entered, so the "
        "summary can tell a refusal from a scope"
    )
    assert first > 0 and numbers  # the parse produced a real body


def test_40_the_summary_accounts_for_every_declared_scenario() -> None:
    """One row per declared scenario, whatever happened to it.

    Four outcomes, four rows: not selected, not entered, entered but not
    measured, measured. A scenario that produced no row would vanish from the
    summary exactly as B and C vanished from runtime run 1.
    """
    source = _executable(TIMING)
    rows = re.findall(r"\$measurements\.Add\(\[pscustomobject\]@\{", source)
    assert len(rows) == 4, (
        f"there must be exactly four measurement-row sites - one per outcome; "
        f"found {len(rows)}"
    )
    assert len(re.findall(r"Selected = \$false", source)) == 1, (
        "exactly one outcome - not selected - may record Selected = $false"
    )
    assert len(re.findall(r"Selected = \$true", source)) == 3, (
        "the other three outcomes must record Selected = $true"
    )
    for branch in ("NOT SELECTED - ", "NOT ENTERED - ", "NOT MEASURED - "):
        assert branch in source, f"the summary must be able to print {branch!r}"
    assert "declared scenario(s) accounted for above" in source, (
        "the summary must state how many of the declared scenarios it accounted "
        "for, so a missing row is visible rather than merely absent"
    )


def test_41_the_scope_of_the_run_is_reported_before_anything_is_measured() -> None:
    """The report must say what it was asked to do, not only what it did.

    Runtime run 1 was a correct scoped run whose report could not be told apart
    from a truncated one. The header now names the scope and lists every
    declared scenario against it.
    """
    source = _text(TIMING)
    assert "SCENARIOS DECLARED, AND THE SCOPE OF THIS RUN" in source, (
        "the report must carry a scope section"
    )
    assert "Write-Phase7Line ('-Scenario ' + $Scenario)" in source, (
        "the scope section must name the -Scenario value it was given"
    )
    assert "'NOT SELECTED for this run'" in source, (
        "the scope section must mark each declared scenario as selected or not"
    )
    scope = source.index("SCENARIOS DECLARED, AND THE SCOPE OF THIS RUN")
    opening = source.index("New-Object -ComObject Excel.Application")
    assert scope < opening, (
        "the scope must be reported before Excel is started, so it survives a "
        "run that fails on its first scenario"
    )
