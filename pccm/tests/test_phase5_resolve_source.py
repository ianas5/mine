#!/usr/bin/env python3
"""PCCM Phase 5 Gate A Step 5: STATIC tests over the resolution layer.

NO VBA IS EXECUTED HERE, AND NONE CAN BE. Every assertion is a statement about
SOURCE TEXT: which procedures exist, in what order they call one another, which
constructs appear in executable code, and which authority each identity comes
from.

Nothing here establishes that the resolver reads a real workbook correctly, that
Excel returns what these procedures expect, or that any resolved number is
right. Those are Gate B's, on real Excel on Windows.

What this file DOES establish:

  * the resolver is the ONLY module allowed workbook access, and the three
    numerical modules are still worksheet-free
  * the reference sets are built from the drivers BEFORE FX or inflation is
    consulted - the ordering that IS the referenced-only rule
  * the reporting currency is a global invariant AND is not seeded into the
    resolved set
  * profiling is looked up by Permanent ID, never by row position
  * identifiers are never trimmed, case-folded or defaulted on the lookup path
  * blanks are refused rather than fabricated as zero
  * the empty driver set is not refused
  * no numerical formula owned by Step 4 is reimplemented here
  * nothing is written to _Calc and no Phase-5 endpoint exists

Runs standalone or under pytest.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder.vba_source import (  # noqa: E402
    VbaModule,
    load_modules,
    logical_statements,
)

SRC_VBA = PCCM_ROOT / "src" / "vba"
RESOLVER = "modCalcResolve"
KERNEL_MODULES = ("modCalcFactors", "modCalcAnalytical", "modCalcFingerprint")

# The resolver's public surface, exact in both directions.
RESOLVER_PUBLIC = {
    "ResolveModel",
    "ResolveAppliedTimeline",
    "ResolveProjectYears",
    "ResolveDrivers",
    "ReferencedCurrencies",
    "ReferencedProfiles",
    "ResolveFxRates",
    "ResolveInflationRates",
    "ResolveInflationFactors",
    "ResolveProfileWeights",
}

# Excel types the resolver may name, because reading a workbook needs them.
WORKBOOK_TYPES = {"ListObject", "Range", "Worksheet"}

# ...and the ones that must never cross its output boundary into a carrier.
CARRIER_FORBIDDEN_TYPES = ("Range", "Worksheet", "ListObject", "Object", "Variant")


def _modules() -> dict[str, VbaModule]:
    return {m.name: m for m in load_modules([SRC_VBA])}


def _resolver() -> VbaModule:
    return _modules()[RESOLVER]


def _synthetic(name: str, body: str) -> VbaModule:
    """A module built from text, for the negative controls. Nothing is executed."""
    return VbaModule(name=name, path=SRC_VBA / f"{name}.bas", raw=body)


def _statements(module: VbaModule, procedure: str) -> list[str]:
    return [text for _, text in logical_statements(_body(module, procedure))]


def _body(module: VbaModule, procedure: str) -> str:
    lines = module.code.splitlines()
    start = next(
        i for i, line in enumerate(lines)
        if re.match(rf"^\s*(Public |Private )?(Static )?(Sub|Function)\s+{procedure}\b", line)
    )
    end = next(i for i in range(start + 1, len(lines))
               if re.match(r"^End (Sub|Function)", lines[i]))
    return "\n".join(lines[start:end])


def _body_raw(module: VbaModule, procedure: str) -> str:
    """The body with string literals intact, for literal-text checks."""
    lines = module.code_without_string_removal.splitlines()
    start = next(
        i for i, line in enumerate(lines)
        if re.match(rf"^\s*(Public |Private )?(Static )?(Sub|Function)\s+{procedure}\b", line)
    )
    end = next(i for i in range(start + 1, len(lines))
               if re.match(r"^End (Sub|Function)", lines[i]))
    return "\n".join(lines[start:end])


def _signature(module: VbaModule, procedure: str) -> str:
    for _, statement in logical_statements(module.code_without_string_removal):
        if re.match(rf"^\s*(Public |Private |Friend )?(Static )?(Sub|Function)\s+{procedure}\b",
                    statement):
            return re.sub(r"\s+", " ", statement)
    raise AssertionError(f"{module.name} does not declare {procedure}")


def call_order(module: VbaModule, procedure: str, names: list[str]) -> list[int]:
    """The statement index of the first call to each name, in the given order.

    `len(statements)` where a name is never called, so a missing call sorts last
    and an out-of-order call is visible as a decreasing index.
    """
    statements = _statements(module, procedure)
    found = []
    for name in names:
        index = next(
            (i for i, s in enumerate(statements) if re.search(rf"(?<![\w.]){name}\s*\(", s)),
            len(statements),
        )
        found.append(index)
    return found


# ===========================================================================
# 1. the module and the boundary
# ===========================================================================
def test_01_the_resolver_exists_and_declares_itself() -> None:
    lines = _resolver().raw.splitlines()
    assert lines[0] == f'Attribute VB_Name = "{RESOLVER}"'
    assert lines[1] == "Option Explicit"


def test_02_the_resolver_is_declared_in_the_structure_contract() -> None:
    import yaml

    contract = yaml.safe_load(
        (PCCM_ROOT / "spec" / "structure_contract.yaml").read_text(encoding="utf-8")
    )
    modules = {m["name"]: m for m in contract["vba"]["modules"]}
    assert RESOLVER in modules, "the resolver must be declared, or the inventory test fails"
    assert modules[RESOLVER]["generated"] is False, (
        "the resolver is hand-written; only modConstants and modCalcContract are generated"
    )
    generated = [m["name"] for m in contract["vba"]["modules"] if m["generated"]]
    assert sorted(generated) == ["modCalcContract", "modConstants"]


def test_03_the_resolver_appears_in_the_stage_b_manifest() -> None:
    import json

    path = PCCM_ROOT / "build" / "stage_b_manifest.json"
    if not path.is_file():
        return
    text = json.dumps(json.loads(path.read_text(encoding="utf-8")))
    assert RESOLVER in text, "the resolver is missing from the Stage-B manifest"


def test_04_the_three_numerical_modules_are_still_worksheet_free() -> None:
    """Allowing the RESOLVER workbook access does not relax the kernel's sweep."""
    tokens = (
        "Application.", "ThisWorkbook", "ActiveWorkbook", "Worksheets", "Worksheet",
        "Range", "Cells", "ListObjects", "ListObject", "Names(", "Evaluate",
        "WorksheetFunction", "modWorkbook.",
    )
    for name in KERNEL_MODULES:
        code = _modules()[name].code
        hits = sorted({t for t in tokens if t.lower() in code.lower()})
        assert hits == [], f"{name} reaches the workbook: {hits}"


def test_05_the_resolver_is_the_only_module_that_reads_the_workbook() -> None:
    resolver = _resolver().code
    assert "modWorkbook." in resolver, (
        "the resolver must go through the Phase-4 access primitives rather than "
        "reaching for ThisWorkbook itself"
    )
    for name in KERNEL_MODULES:
        assert "modCalcResolve." not in _modules()[name].code, (
            f"{name} calls the resolver; the dependency runs the other way"
        )


def test_06_no_excel_object_crosses_the_output_boundary() -> None:
    """A carrier field may not be a Range, a Worksheet, a ListObject or untyped."""
    raw = _resolver().raw
    for type_name in ("ResolvedTimeline", "ResolvedDriver", "ResolvedModel"):
        block = re.search(rf"^Public Type {type_name}$(.*?)^End Type$", raw,
                          re.MULTILINE | re.DOTALL)
        assert block, f"{type_name} is not declared as a Public Type"
        for forbidden in CARRIER_FORBIDDEN_TYPES:
            assert not re.search(rf"\bAs\s+{forbidden}\b", block.group(1)), (
                f"{type_name} exposes {forbidden}"
            )


def test_07_a_workbook_row_number_is_never_a_driver_identity() -> None:
    """Identity is the Permanent ID. A row index is a way to reach a cell."""
    raw = _resolver().raw
    block = re.search(r"^Public Type ResolvedDriver$(.*?)^End Type$", raw,
                      re.MULTILINE | re.DOTALL)
    assert "PermanentId As String" in block.group(1)
    assert not re.search(r"\b(Row|RowIndex|WorksheetRow)\s+As\s+Long", block.group(1)), (
        "the driver carrier stores a workbook row as identity"
    )


def test_08_no_phase_5_endpoint_or_later_module_was_added() -> None:
    modules = _modules()
    executable = "\n".join(m.code for m in modules.values())
    declared = {p for m in modules.values() for p in m.procedures}
    for deferred in ("modCalcCheck", "modCalcReport", "PCCM_Calculate",
                     "PCCM_CalculationStatus", "PCCM_CalculationAttemptResult",
                     "PCCM_CalculationAttemptDetail", "PCCM_CalculationFingerprint",
                     "PCCM_CurrentInputFingerprint"):
        assert deferred not in executable
        assert deferred not in declared
        assert deferred not in modules
    assert [p for p in _resolver().procedures if p.startswith("PCCM_")] == []


def test_09_nothing_is_written_anywhere() -> None:
    """Step 5 resolves. It does not write, and it never touches _Calc."""
    code = _resolver().code
    for writer in ("modWorkbook.WriteValue", ".ClearContents", "ListRows.Add",
                   "ListColumns.Add", ".Delete", "SH_CALC", "CALC_SHEET"):
        assert writer not in code, f"the resolver performs a write or reaches _Calc ({writer})"
    assert not re.search(r"\.Value\s*=[^=]", code), "the resolver assigns a cell value"


# ===========================================================================
# 2. the ordering rule
# ===========================================================================
def test_10_reference_sets_are_built_before_fx_and_inflation_are_consulted() -> None:
    """THE referenced-only rule, as an order of operations.

    A Config assumption for a currency or profile nobody uses cannot block a
    valid model, because resolution never asks about it. That is only true if
    the drivers are identified and the reference sets derived first.
    """
    order = call_order(_resolver(), "ResolveModel", [
        "ResolveDrivers", "ReferencedCurrencies", "ReferencedProfiles",
        "ResolveFxRates", "ResolveInflationRates",
    ])
    drivers, currencies, profiles, fx, inflation = order
    assert drivers < currencies < fx, "FX is consulted before the reference set exists"
    assert drivers < profiles < inflation, (
        "inflation is consulted before the reference set exists"
    )
    assert currencies < fx and profiles < inflation


def test_11_the_reference_sets_are_derived_from_the_drivers_alone() -> None:
    module = _resolver()
    for procedure, field in (("ReferencedCurrencies", "Currency"),
                             ("ReferencedProfiles", "InflationProfile")):
        body = _body(module, procedure)
        assert f".{field}" in body, f"{procedure} does not read the driver field"
        for table in ("TBL_FX_RATES", "TBL_INFLATION", "modWorkbook."):
            assert table not in body, (
                f"{procedure} consults the workbook; a reference set is derived from "
                "the drivers, not discovered in Config"
            )


def test_12_fx_lookup_is_constrained_to_the_referenced_set() -> None:
    body = _body(_resolver(), "ResolveFxRates")
    assert "For index = 0 To nameCount - 1" in body, (
        "the FX loop must walk the referenced names, not the FX table"
    )
    assert not re.search(r"For \w+ = 1 To modWorkbook\.BodyRowCount", body), (
        "the resolver walks every FX row; an unreferenced bad row would then block"
    )


def test_13_inflation_lookup_is_constrained_to_the_referenced_profiles() -> None:
    body = _body(_resolver(), "ResolveInflationRates")
    assert "For index = 0 To nameCount - 1" in body
    assert not re.search(r"For \w+ = 1 To modWorkbook\.BodyRowCount", body), (
        "the resolver walks every inflation row; an unreferenced incomplete "
        "profile would then block"
    )
    assert "MatchingGridRow(table, GCOL_INFLATION_PROFILE_NAME, key)" in body


# ===========================================================================
# 3. the reporting currency
# ===========================================================================
def test_14_the_reporting_currency_has_an_explicit_global_invariant() -> None:
    module = _resolver()
    body = _body(module, "ResolveFxRates")
    statements = _statements(module, "ResolveFxRates")
    invariant = next(i for i, s in enumerate(statements)
                     if "ReportingCurrencyInvariant(table, detail)" in s)
    empty = next(i for i, s in enumerate(statements) if s == "If nameCount = 0 Then")
    assert invariant < empty, (
        "the invariant must hold even for a model that references no currency at all"
    )
    guard = _body(module, "ReportingCurrencyInvariant")
    assert "MatchingFxRows(table, REPORTING_CURRENCY, row)" in guard
    assert "If matches <> 1 Then" in guard, "exactly one row must carry it"
    assert "If rate <> REPORTING_CURRENCY_RATE Then" in guard, "and it must equal the identity"
    assert "REPORTING_CURRENCY" in body


def test_15_the_resolved_set_is_not_seeded_with_the_reporting_currency() -> None:
    """Being validated globally does not make a currency referenced.

    A USD-only model resolves USD and nothing else; an empty model resolves
    nothing. The resolved rows are the referenced ones.
    """
    body = _body(_resolver(), "ResolveFxRates")
    assert "ReDim rates(0 To nameCount - 1)" in body, (
        "the resolved set is sized from the REFERENCED names"
    )
    assert not re.search(r"nameCount\s*\+\s*1", body), (
        "the resolved set is widened beyond the referenced names"
    )
    statements = _statements(_resolver(), "ResolveFxRates")
    empty = next(i for i, s in enumerate(statements) if s == "If nameCount = 0 Then")
    assert statements[empty + 1] == "ResolveFxRates = True", (
        "an empty reference set must resolve to an empty rate set, not to one row"
    )


def test_16_the_reporting_currency_and_its_rate_are_projected_not_restated() -> None:
    """"SAR" and its identity rate belong to the FX table's own locked seed row."""
    module = _resolver()
    assert '"SAR"' not in module.code_without_string_removal, (
        "the reporting currency is restated instead of projected"
    )
    assert "REPORTING_CURRENCY" in module.code
    assert "REPORTING_CURRENCY_RATE" in module.code
    for name in ("REPORTING_CURRENCY", "REPORTING_CURRENCY_RATE",
                 "COL_FX_RATES_CURRENCY", "COL_FX_RATES_FX_TO_SAR",
                 "NM_INPUT_DISCOUNT_RATE", "DISTRIBUTION_NAME_1"):
        assert name in _projected_constants(), f"{name} is not projected by the builder"


def _projected_constants() -> set[str]:
    """The constants the EMITTER produces, rendered in process.

    Read from the generator rather than from `build/`, so the assertion is about
    the authority and not about whether a build artifact happens to be on disk.
    """
    from pccm_builder import (
        load_contract, load_driver_contract, load_spec, load_structure_contract,
    )
    from pccm_builder.stage_b_emit import render_constants_module

    spec_dir = PCCM_ROOT / "spec"
    text = render_constants_module(
        load_spec(spec_dir / "workbook.yaml"),
        load_contract(spec_dir / "input_contract.yaml"),
        load_driver_contract(spec_dir / "driver_contract.yaml"),
        load_structure_contract(spec_dir / "structure_contract.yaml"),
    )
    return {
        match.group(1)
        for match in re.finditer(r"^Public Const (\w+)", text, re.MULTILINE)
    }


def test_16a_the_projected_values_come_from_the_contract_not_the_emitter() -> None:
    """The reporting currency is the FX table's own locked seed row.

    Writing "SAR" into the emitter would move the authority; reading the seed row
    keeps it where the contract already put it, so a change to that row is a
    change to the projection.
    """
    from pccm_builder import load_contract
    from pccm_builder.stage_b_emit import render_constants_module

    contract = load_contract(PCCM_ROOT / "spec" / "input_contract.yaml")
    fx = next(t for t in contract.all_tables if t.table_name == "tblFXRates")
    assert fx.locked_seed_rows == 1 and len(fx.seed_rows) == 1
    currency, rate = fx.seed_rows[0][0], fx.seed_rows[0][1]
    source = (PCCM_ROOT / "builder" / "pccm_builder" / "stage_b_emit.py").read_text(
        encoding="utf-8"
    )
    assert f'"{currency}"' not in source, "the reporting currency is restated in the emitter"
    assert f'Public Const REPORTING_CURRENCY As String = "{currency}"' in _rendered()
    assert f"Public Const REPORTING_CURRENCY_RATE As Double = {float(rate)}" in _rendered()


def _rendered() -> str:
    from pccm_builder import (
        load_contract, load_driver_contract, load_spec, load_structure_contract,
    )
    from pccm_builder.stage_b_emit import render_constants_module

    spec_dir = PCCM_ROOT / "spec"
    return render_constants_module(
        load_spec(spec_dir / "workbook.yaml"),
        load_contract(spec_dir / "input_contract.yaml"),
        load_driver_contract(spec_dir / "driver_contract.yaml"),
        load_structure_contract(spec_dir / "structure_contract.yaml"),
    )


def test_17_a_duplicate_or_missing_referenced_rate_is_a_failure() -> None:
    module = _resolver()
    counter = _body(module, "MatchingFxRows")
    assert "found = found + 1" in counter, (
        "matches must be COUNTED, so a duplicate is reported rather than resolved "
        "by taking the first row"
    )
    body = _body(module, "ResolveFxRates")
    assert "If matches <> 1 Then" in body
    assert "If rate <= 0# Then" in body, "a referenced rate must be strictly positive"


# ===========================================================================
# 4. calendar-year anchoring and the applied timeline
# ===========================================================================
def test_18_the_resolver_reads_applied_values_and_never_entered_ones() -> None:
    module = _resolver()
    code = module.code
    for applied in ("NM_APPLIED_BASE_YEAR", "NM_APPLIED_START_YEAR",
                    "NM_APPLIED_DURATION", "NM_APPLIED_LAST_YEAR"):
        assert applied in code, f"{applied} is never read"
    for entered in ("NM_BASE_YEAR_ENTERED", "NM_PROJECT_START_YEAR_ENTERED",
                    "NM_DURATION_YEARS_ENTERED"):
        assert entered not in code, (
            f"{entered} is read; an entered value has not generated its columns"
        )


def test_19_the_required_inflation_span_is_anchored_to_calendar_years() -> None:
    """`BaseYear + 1 .. LastProjectYear`, selected by year and not by position."""
    body = _body(_resolver(), "ResolveInflationRates")
    assert "yearCount = timeline.LastYear - timeline.BaseYear" in body
    assert "year = timeline.BaseYear + 1 + offset" in body
    assert "column = YearColumn(table, GRID_INFLATION_FIXED_COLS, year)" in body, (
        "the column must be located by its calendar-year header"
    )


def test_20_a_year_column_is_located_by_its_header_not_by_arithmetic() -> None:
    """A grid not regenerated for the applied timeline must report a missing
    column rather than silently read the wrong year."""
    body = _body(_resolver(), "YearColumn")
    assert "table.HeaderRowRange.Cells(1, columnIndex)" in body
    assert "If value = CDbl(headerValue) Then" in body


def test_21_no_base_year_rate_is_invented() -> None:
    """The Base Year has no annual rate; the span starts at BaseYear + 1."""
    body = _body(_resolver(), "ResolveInflationRates")
    assert "timeline.BaseYear + 1" in body
    assert not re.search(r"rates\(index, .*\) = 0#", body), (
        "a Base-Year rate of zero is fabricated to fill the array"
    )


def test_22_the_project_year_map_starts_at_the_start_year() -> None:
    body = _body(_resolver(), "ResolveProjectYears")
    assert "calendarYears(index) = timeline.StartYear + index" in body
    assert "projectIndexes(index) = index + 1" in body


# ===========================================================================
# 5. profiling
# ===========================================================================
def test_23_profiling_is_resolved_by_permanent_id() -> None:
    """A driver reorder must not attach another driver's weights."""
    body = _body(_resolver(), "ResolveProfileWeights")
    assert "MatchingGridRow(grid, keyColumn, drivers(LBound(drivers) + index).PermanentId)" in body
    assert "If row = 0 Then" in body, "an unmatched Permanent ID must be a failure"
    lookup = _body(_resolver(), "MatchingGridRow")
    assert "vbBinaryCompare" in lookup


def test_24_profiling_never_walks_the_two_tables_in_parallel() -> None:
    body = _body(_resolver(), "ResolveProfileWeights")
    assert not re.search(r"CellIn\(grid, index", body), (
        "the grid row is taken from the driver's position in the register"
    )
    assert not re.search(r"row = index", body), "the grid row is the register row"


def test_25_only_applied_project_year_columns_are_read() -> None:
    body = _body(_resolver(), "ResolveProfileWeights")
    assert "For offset = 0 To timeline.Duration - 1" in body
    assert "column = YearColumn(grid, fixedCols, offset + 1)" in body


def test_26_a_blank_weight_is_refused_and_never_becomes_zero() -> None:
    module = _resolver()
    reader = _body_raw(module, "NumericCell")
    assert "modWorkbook.IsEmptyCell(cell)" in reader
    assert "A blank is not zero." in reader
    statements = _statements(module, "NumericCell")
    blank = statements.index("If modWorkbook.IsEmptyCell(cell) Then")
    read = next(i for i, s in enumerate(statements) if "TryReadDouble" in s)
    assert blank < read, "the blank must be caught before any numeric coercion"
    assert not re.search(r"weight = 0#", _body(module, "ResolveProfileWeights")), (
        "a blank profiling cell is fabricated as zero"
    )


# ===========================================================================
# 6. identifiers
# ===========================================================================
def test_27_identifiers_are_read_raw_and_never_trimmed_into_another_key() -> None:
    module = _resolver()
    raw_reader = _body(module, "RawCellText")
    assert "text = CStr(cell.Value)" in raw_reader
    assert "Trim$" not in raw_reader, "the raw reader trims; a key would be rewritten"
    assert "LCase" not in raw_reader and "UCase" not in raw_reader


def test_28_the_identifier_gate_refuses_rather_than_repairs() -> None:
    module = _resolver()
    body = _body(module, "ExactIdentifier")
    assert "value = text" in body, "the identifier is stored exactly as read"
    assert not re.search(r"value = Trim\$\(", body), "the identifier is trimmed into the model"
    # Whitespace-only is REFUSED. Trim$ appears only in that test, never in the
    # value that is kept.
    assert "If Len(Trim$(text)) = 0 Then" in body
    assert "ExactIdentifier = True" in body


def test_29_no_lookup_compares_case_insensitively_or_on_a_repaired_key() -> None:
    code = _resolver().code
    assert "vbTextCompare" not in code, "a case-insensitive lookup would repair a key"
    for lookup in ("MatchingFxRows", "MatchingGridRow"):
        body = _body(_resolver(), lookup)
        assert "vbBinaryCompare" in body
        assert "Trim$" not in body and "LCase" not in body and "UCase" not in body


def test_30_an_unknown_distribution_is_refused_not_defaulted() -> None:
    module = _resolver()
    adapter = _body(module, "DistributionKindOf")
    for constant in ("DISTRIBUTION_NAME_1", "DISTRIBUTION_NAME_2", "DISTRIBUTION_NAME_3"):
        assert constant in adapter, f"{constant} is not part of the adapter"
    assert "Case Else" not in adapter, "an unknown distribution is mapped to a default"
    caller = _body(module, "ReadDriverRow")
    assert "If driver.DistKind = 0 Then" in caller, (
        "an unmapped distribution must be a controlled failure"
    )


# ===========================================================================
# 7. the empty driver set
# ===========================================================================
def test_31_the_empty_driver_set_is_not_refused() -> None:
    """No accepted contract requires at least one Cost Line or Risk."""
    module = _resolver()
    body = _body(module, "ResolveDrivers")
    statements = _statements(module, "ResolveDrivers")
    guard = statements.index("If capacity < 1 Then")
    assert statements[guard + 1] == "ResolveDrivers = True"
    assert "driverCount = 0" in body
    for procedure, count in (("ReferencedCurrencies", "driverCount"),
                             ("ReferencedProfiles", "driverCount"),
                             ("ResolveProfileWeights", "driverCount"),
                             ("ResolveFxRates", "nameCount")):
        statements = _statements(module, procedure)
        empty = statements.index(f"If {count} = 0 Then")
        assert statements[empty + 1] == f"{procedure} = True", (
            f"{procedure} does not succeed on an empty set"
        )


def test_32_the_empty_branches_precede_any_array_access() -> None:
    """A VBA array cannot represent a zero-element set, so the count is explicit
    and is tested before any bound is read - the same rule the Step-4 aggregate
    boundaries follow."""
    module = _resolver()
    for procedure, count, array in (("ReferencedCurrencies", "driverCount", "drivers"),
                                    ("ReferencedProfiles", "driverCount", "drivers"),
                                    ("ResolveProfileWeights", "driverCount", "drivers")):
        statements = _statements(module, procedure)
        empty = statements.index(f"If {count} = 0 Then")
        touch = next(
            (i for i, s in enumerate(statements)
             if i and re.search(rf"[LU]Bound\(\s*{array}\b", s)),
            len(statements),
        )
        assert empty < touch, f"{procedure} reads a bound of {array} before the empty branch"
        assert f"If {count} < 0 Then Exit Function" in statements, (
            f"{procedure} must refuse a negative count"
        )


def test_33_the_global_invariant_still_runs_for_an_empty_model() -> None:
    """An empty driver set does not excuse a broken reporting currency."""
    statements = _statements(_resolver(), "ResolveFxRates")
    invariant = next(i for i, s in enumerate(statements)
                     if "ReportingCurrencyInvariant" in s)
    empty = statements.index("If nameCount = 0 Then")
    assert invariant < empty


# ===========================================================================
# 8. no second implementation of an accepted formula
# ===========================================================================
def test_34_the_resolver_reimplements_no_step_4_formula() -> None:
    """The adapter layer adapts. It does not compute.

    Every one of these is owned by an accepted numerical module, and a second
    implementation is two chances to disagree.
    """
    code = _resolver().code
    for owned in ("TriangularMean", "PertMean", "UniformMean", "DeterministicCentral",
                  "ExpectedRisk", "BuildKnom", "BuildKpv", "AccumulateTotals",
                  "BuildAnnualSeries", "Reconcile", "CalcFpDigestStream",
                  "CalcFpBuildFingerprint", "ExactSumOfProducts"):
        assert owned not in code, f"the resolver reaches into {owned}"
    # No compounding, no discounting, no distribution arithmetic of its own.
    assert not re.search(r"running\s*=\s*running\s*\*", code), "the resolver compounds"
    assert not re.search(r"/\s*3#|/\s*6#|\*\s*2#\s*/\s*3#", code), (
        "the resolver computes a distribution statistic"
    )


def test_35_a_resolved_factor_is_produced_by_the_accepted_function() -> None:
    body = _body(_resolver(), "ResolveInflationFactors")
    assert "modCalcFactors.BuildInflationFactors(" in body, (
        "the cumulative factors must come from the accepted builder"
    )
    assert not re.search(r"factors\(\w+\)\s*=\s*\w+\s*\*", body), (
        "the resolver builds the series itself"
    )


def test_36_range_checking_goes_through_the_accepted_predicate() -> None:
    for procedure in ("NumericCell", "NumericNamedCell"):
        assert "modCalcFactors.IsUsableDouble(value)" in _body(_resolver(), procedure)


# ===========================================================================
# 9. error handling and the public surface
# ===========================================================================
def test_37_no_generic_error_suppression() -> None:
    code = _resolver().code
    assert "On Error Resume Next" not in code, (
        "a suppressed workbook error becomes valid model data"
    )
    assert "On Error" not in code, (
        "the resolver installs no handler at all; every failure is a returned False"
    )


def test_38_every_failure_carries_a_detail() -> None:
    """A resolution failure a user cannot act on is barely better than a crash."""
    module = _resolver()
    for procedure in sorted(RESOLVER_PUBLIC):
        assert "detail" in _signature(module, procedure), (
            f"{procedure} cannot report why it failed"
        )


def test_39_the_public_surface_is_exactly_the_whitelist() -> None:
    actual = set(_resolver().public_procedures)
    assert actual == RESOLVER_PUBLIC, (
        f"unexpected: {sorted(actual - RESOLVER_PUBLIC)}; "
        f"missing: {sorted(RESOLVER_PUBLIC - actual)}"
    )


def test_40_the_workbook_types_stay_inside_the_module() -> None:
    """`ListObject` is a parameter of PRIVATE helpers only."""
    module = _resolver()
    for procedure in sorted(RESOLVER_PUBLIC):
        signature = _signature(module, procedure)
        for type_name in WORKBOOK_TYPES:
            assert not re.search(rf"\bAs\s+{type_name}\b", signature), (
                f"{procedure} exposes {type_name} across the public boundary"
            )


# ===========================================================================
# 10. NEGATIVE CONTROLS
#
# Each plants the regression the rule exists to prevent and asserts the sweep
# that would catch it does.
# ===========================================================================
_STUB = 'Attribute VB_Name = "modProbe"\nOption Explicit\n'


def test_nc_01_validating_every_fx_row_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function ResolveFxRates() As Boolean\n"
        "    For rowIndex = 1 To modWorkbook.BodyRowCount(table)\n"
        "        If Not NumericCell(table, rowIndex, COL_FX_RATES_FX_TO_SAR, rate) Then\n"
        "            Exit Function\n        End If\n    Next rowIndex\n"
        "End Function\n",
    )
    body = _body(planted, "ResolveFxRates")
    assert re.search(r"For \w+ = 1 To modWorkbook\.BodyRowCount", body), (
        "walking every FX row must be visible to the sweep"
    )


def test_nc_02_seeding_the_resolved_set_with_the_reporting_currency_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function ResolveFxRates() As Boolean\n"
        "    ReDim rates(0 To nameCount + 1)\n"
        "    rates(0) = REPORTING_CURRENCY_RATE\n"
        "End Function\n",
    )
    body = _body(planted, "ResolveFxRates")
    assert re.search(r"nameCount\s*\+\s*1", body), "the widened set must be visible"
    assert "ReDim rates(0 To nameCount - 1)" not in body


def test_nc_03_resolving_every_inflation_profile_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function ResolveInflationRates() As Boolean\n"
        "    For rowIndex = 1 To modWorkbook.BodyRowCount(table)\n"
        "        rates(rowIndex) = 0#\n    Next rowIndex\n"
        "End Function\n",
    )
    body = _body(planted, "ResolveInflationRates")
    assert re.search(r"For \w+ = 1 To modWorkbook\.BodyRowCount", body)
    assert "For index = 0 To nameCount - 1" not in body


def test_nc_04_attaching_weights_by_row_position_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function ResolveProfileWeights() As Boolean\n"
        "    row = index\n"
        "    weight = modWorkbook.CellIn(grid, index, column).Value\n"
        "End Function\n",
    )
    body = _body(planted, "ResolveProfileWeights")
    assert re.search(r"row = index", body), "positional attachment must be visible"
    assert re.search(r"CellIn\(grid, index", body)
    assert "MatchingGridRow(" not in body


def test_nc_05_reading_entered_structural_values_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function ResolveAppliedTimeline() As Boolean\n"
        "    value = modWorkbook.ReadValue(NM_DURATION_YEARS_ENTERED)\n"
        "End Function\n",
    )
    assert "NM_DURATION_YEARS_ENTERED" in planted.code


def test_nc_06_a_blank_weight_becoming_zero_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function NumericCell() As Boolean\n"
        "    If Not modWorkbook.TryReadDouble(cell.Value, value) Then\n"
        "        value = 0#\n        NumericCell = True\n        Exit Function\n"
        "    End If\n"
        "    If modWorkbook.IsEmptyCell(cell) Then Exit Function\n"
        "End Function\n",
    )
    statements = _statements(planted, "NumericCell")
    blank = statements.index("If modWorkbook.IsEmptyCell(cell) Then Exit Function")
    read = next(i for i, s in enumerate(statements) if "TryReadDouble" in s)
    assert read < blank, "the fabricated zero must be visible as an ordering defect"
    assert "value = 0#" in _body(planted, "NumericCell")


def test_nc_07_trimming_an_identifier_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function ExactIdentifier() As Boolean\n"
        "    value = Trim$(text)\n    ExactIdentifier = True\n"
        "End Function\n",
    )
    body = _body(planted, "ExactIdentifier")
    assert re.search(r"value = Trim\$\(", body), "the repaired key must be visible"


def test_nc_08_a_case_insensitive_lookup_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function MatchingGridRow() As Long\n"
        "    If StrComp(text, key, vbTextCompare) = 0 Then MatchingGridRow = rowIndex\n"
        "End Function\n",
    )
    assert "vbTextCompare" in planted.code


def test_nc_09_refusing_an_empty_driver_model_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function ResolveDrivers() As Boolean\n"
        "    If capacity < 1 Then\n        detail = \"at least one driver is required\"\n"
        "        Exit Function\n    End If\n"
        "End Function\n",
    )
    statements = _statements(planted, "ResolveDrivers")
    guard = statements.index("If capacity < 1 Then")
    assert statements[guard + 1] != "ResolveDrivers = True", (
        "the planted refusal must not look like the accepted empty path"
    )
    assert "at least one driver is required" in _body_raw(planted, "ResolveDrivers")


def test_nc_10_worksheet_access_in_a_numerical_module_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function BuildKnom() As Boolean\n"
        "    BuildKnom = ThisWorkbook.Worksheets(\"Setup\").Range(\"A1\").Value\n"
        "End Function\n",
    )
    tokens = ("ThisWorkbook", "Worksheets", "Range")
    hits = sorted({t for t in tokens if t.lower() in planted.code.lower()})
    assert hits == ["Range", "ThisWorkbook", "Worksheets"]


def test_nc_11_an_early_calculation_endpoint_is_caught() -> None:
    planted = _synthetic("modProbe", _STUB + "Public Sub PCCM_Calculate()\nEnd Sub\n")
    assert "PCCM_Calculate" in planted.procedures


def test_nc_12_a_calc_writeback_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub Publish()\n"
        "    modWorkbook.Sh(SH_CALC).Range(CALC_STATE_VALUE_RANGE).Value = 1\n"
        "End Sub\n",
    )
    code = planted.code
    assert "SH_CALC" in code
    assert re.search(r"\.Value\s*=[^=]", code), "the write must be visible to the sweep"


def test_nc_13_reimplementing_an_accepted_formula_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function Compound() As Boolean\n"
        "    running = running * growth\n"
        "End Function\n",
    )
    assert re.search(r"running\s*=\s*running\s*\*", planted.code)


def test_nc_14_generic_error_suppression_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function ResolveModel() As Boolean\n"
        "    On Error Resume Next\n    ResolveModel = True\n"
        "End Function\n",
    )
    assert "On Error Resume Next" in planted.code


# ===========================================================================
# 11. this suite makes no runtime claim
# ===========================================================================
def test_41_no_test_in_this_file_claims_that_vba_ran() -> None:
    text = Path(__file__).read_text(encoding="utf-8")
    banned = (
        ("VBA", "produced"), ("VBA", "computed"), ("VBA", "returned"),
        ("VBA", "evaluated"), ("read", "the real workbook"),
        ("resolved", "at runtime"), ("executed", "the VBA"), ("ran", "the VBA"),
    )
    for parts in banned:
        assert " ".join(parts) not in text, f"this suite must not make that claim: {parts}"
    assert "NO VBA IS EXECUTED HERE" in text


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
