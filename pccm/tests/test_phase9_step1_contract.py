#!/usr/bin/env python3
"""The Phase-9 Step-1 contract says only things the source can be asked.

WHAT A CONTRACT RECORD IS FOR. It is the authority P9-2 will be built against.
Every claim it makes about an existing owner is therefore re-derived here, so a
plan built on it cannot rest on a fact that has quietly stopped being true.

AND IT IS A CONTRACT, NOT AN IMPLEMENTATION. The last control refuses the day
this record starts describing work that exists: no Phase-9 production, spec or
builder change may be in the tree while this document says none is.

Runs standalone or under pytest.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
RECORD = PCCM_ROOT / "docs" / "phase9_step1_contract.md"
SRC = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402
import yaml  # noqa: E402


def _text() -> str:
    return RECORD.read_text(encoding="utf-8")


def _module(name: str) -> str:
    return (SRC / name).read_text(encoding="utf-8")


def _procedure(module: str, name: str) -> str:
    source = _module(module)
    match = re.search(
        rf"^(?:Public|Private)\s+(?:Function|Sub)\s+{re.escape(name)}\b.*?^End (?:Function|Sub)",
        source, re.S | re.M)
    assert match, f"{name} is not defined in {module}"
    return match.group(0)


# ===========================================================================
# A. IT IS A CONTRACT, AND NOTHING HAS BEEN BUILT FROM IT YET
# ===========================================================================
def test_01_the_record_declares_itself_design_only() -> None:
    text = _text()
    assert RECORD.is_file()
    assert "CONTRACT / DESIGN ONLY." in text
    assert "P9-2 HAS NOT STARTED." in text


def test_02_no_phase_9_implementation_exists_in_the_tree() -> None:
    """THE CLAIM NOBODY WOULD OTHERWISE CHECK."""
    strays = []
    for directory, pattern in ((PCCM_ROOT / "builder" / "pccm_builder", "phase9*"),
                               (PCCM_ROOT / "bootstrap" / "windows", "*phase9*"),
                               (PCCM_ROOT / "build", "phase9*")):
        strays += [p.name for p in directory.glob(pattern)]
    assert not strays, f"Phase-9 implementation exists while the record says none does: {strays}"
    # AND THE TWO SPEC ADDITIONS IT PLANS ARE NOT THERE YET.
    manifest = yaml.safe_load((SPEC / "workbook.yaml").read_text(encoding="utf-8"))
    assert "phase9_shell" not in manifest, "phase9_shell already exists"
    inputs = yaml.safe_load((SPEC / "input_contract.yaml").read_text(encoding="utf-8"))
    assert "recommended_iterations" not in inputs["inputs"]["monte_carlo_iterations"], (
        "the recommendation field already exists; §9.1 is out of date")


# ===========================================================================
# B. THE OWNERS IT NAMES ARE REAL
# ===========================================================================
@pytest.mark.parametrize("module,name", [
    ("modStructuralCheck.bas", "ValidateStructure"),
    ("modStructuralCheck.bas", "PCCM_StructuralReport"),
    ("modCalcReport.bas", "DeriveStatus"),
    ("modCalcReport.bas", "PrepareCurrentCalculation"),
    ("modCalcReport.bas", "PCCM_CalculationStatus"),
    ("modSimReport.bas", "SimReportDerivedStatus"),
    ("modSimReport.bas", "DeriveSimStatus"),
    ("modResultsState.bas", "PCCM_ResultsSimulationState"),
    ("modSimAnnualStore.bas", "PCCM_AnnualDistributionState"),
])
def test_10_every_named_owner_exists(module: str, name: str) -> None:
    assert _procedure(module, name)
    assert name in _text(), f"the record does not name {name}"


def test_11_the_adapter_it_authorises_does_not_exist_yet() -> None:
    """THE AUTHORISATION IS FOR WORK NOT DONE."""
    for name in ("CalcReportDerivedStatus", "PCCM_ModelCheckCalculationState"):
        assert name in _text(), f"the record does not name {name}"
        assert name not in _module("modCalcReport.bas"), f"{name} already exists"
        assert name not in _module("modResultsState.bas"), f"{name} already exists"


# ===========================================================================
# C. THE WORKSHEET-SAFETY AUDIT IS RE-DERIVED, NOT ASSERTED
# ===========================================================================
WRITES = (r"\.Value2\s*=", r"\.Value\s*=", r"\.Formula[0-9A-Za-z]*\s*=",
          r"ClearContents", r"\.Delete\b", r"ListRows\.Add",
          r"EnableEvents\s*=", r"ScreenUpdating\s*=", r"\.Calculation\s*=")


def _writes_in(text: str) -> list[str]:
    found = []
    for pattern in WRITES:
        found += [m.group(0) for m in re.finditer(pattern, text)]
    return found


def test_20_the_structural_module_really_writes_nothing() -> None:
    """THE RECORD SAYS `PCCM_StructuralReport` IS SAFE AS IT STANDS."""
    writes = _writes_in(_module("modStructuralCheck.bas"))
    assert not writes, f"modStructuralCheck writes: {sorted(set(writes))}"
    assert "**SAFE as it stands.** No split needed" in _text()


def test_21_the_calculation_status_really_does_write() -> None:
    """THE HAZARD THE ADAPTER EXISTS FOR."""
    body = _procedure("modCalcReport.bas", "PCCM_CalculationStatus")
    assert "WriteStatusBlock" in body, (
        "PCCM_CalculationStatus no longer persists; §8's premise is gone")
    assert _writes_in(_procedure("modCalcReport.bas", "WriteStatusBlock"))
    assert "writes `C19:C20`" in _text()


def test_22_the_derivation_it_will_expose_is_pure() -> None:
    """THE OTHER HALF OF §8: the split is safe because the derivation already is."""
    for name in ("DeriveStatus", "PrepareCurrentCalculation"):
        writes = _writes_in(_procedure("modCalcReport.bas", name))
        assert not writes, f"{name} writes: {sorted(set(writes))}"


def test_23_the_phase_8_read_paths_are_still_read_only() -> None:
    for module, name in (("modSimReport.bas", "SimReportDerivedStatus"),
                         ("modSimReport.bas", "DeriveSimStatus"),
                         ("modResultsState.bas", "PCCM_ResultsSimulationState")):
        writes = _writes_in(_procedure(module, name))
        assert not writes, f"{name} writes: {sorted(set(writes))}"


# ===========================================================================
# D. THE TWO-AXIS AND PRECEDENCE CLAIMS
# ===========================================================================
def test_30_the_calculation_axis_is_the_four_words_it_names() -> None:
    body = _procedure("modCalcReport.bas", "DeriveStatus")
    for constant in ("CALC_STATUS_INVALID", "CALC_STATUS_NOT_CALCULATED",
                     "CALC_STATUS_CURRENT", "CALC_STATUS_STALE"):
        assert constant in body, constant
    assert "`NOT CALCULATED / CURRENT / STALE / INVALID`" in _text()


def test_31_the_simulation_invalid_precedence_is_derivable() -> None:
    """§2.2 RESTS ON TWO SOURCE FACTS, AND BOTH ARE CHECKED HERE."""
    derive = _procedure("modSimReport.bas", "DeriveSimStatus")
    assert "CurrentRequestFingerprint" in derive
    assert "SIM_STATE_INVALID" in derive
    prepare = _procedure("modCalcReport.bas", "CalcPrepareSimulationInputs")
    assert "CALC_STATUS_CURRENT" in prepare, (
        "the simulation no longer requires a CURRENT calculation; §2.2 is void")
    assert "the simulation needs a CURRENT calculation" in prepare
    text = _text()
    assert ("Simulation `INVALID` is an actionable **ERROR** if and only if\n"
            "> the calculation state is `CURRENT`.") in text


def test_32_the_simulation_axis_still_has_a_blank_and_three_words() -> None:
    contract = yaml.safe_load((SPEC / "sim_contract.yaml").read_text(encoding="utf-8"))
    assert contract["label_sets"]["sim_state"] == ["CURRENT", "STALE", "INVALID"]
    assert contract["sim_state"]["no_success_valid_status"] is None
    assert "a **blank** meaning no\npublication exists" in _text()


# ===========================================================================
# E. THE THRESHOLD OWNERSHIP CONCLUSION
# ===========================================================================
def test_40_the_advisory_is_already_assigned_to_model_check() -> None:
    """THE FINDING THAT MAKES PHASE 9 AN OBLIGATION, NOT AN INVENTION."""
    inputs = yaml.safe_load((SPEC / "input_contract.yaml").read_text(encoding="utf-8"))
    note = inputs["inputs"]["monte_carlo_iterations"]["note"]
    assert "the <10000 advisory belongs to Model Check" in note, note
    assert "the <10000\n> advisory belongs to Model Check" in _text()


def test_41_the_hard_minimum_is_unchanged_and_is_not_the_threshold() -> None:
    inputs = yaml.safe_load((SPEC / "input_contract.yaml").read_text(encoding="utf-8"))
    iterations = inputs["inputs"]["monte_carlo_iterations"]
    assert iterations["validation"]["formula1"] == "1000"
    assert iterations["validation"]["operator"] == "greaterThanOrEqual"
    text = _text()
    assert "The hard minimum of 1000 is unchanged and untouched" in text
    # AND THE RECORD REFUSES THE TWO WRONG SOURCES FOR THE THRESHOLD.
    assert "The prompt is **not** parsed" in text
    assert "`default` is **not** the recommendation" in text
    assert iterations["default"] == 10000, (
        "the default moved; §9.1's 'they coincide today' is out of date")


# ===========================================================================
# F. THE DECISIONS THIS STEP WAS ASKED TO SETTLE
# ===========================================================================
def test_50_every_settled_decision_is_stated() -> None:
    text = _text()
    for decision in (
            "row_window = 100",
            "Showing the first 100 of <M> checks.",
            "Calculation `NOT CALCULATED` | **WARNING**",
            "Monte Carlo iterations are below the recommended 10,000; "
            "simulation precision may be lower at this setting.",
            "Increase Monte Carlo Iterations on Setup to 10,000 or more "
            "before the next simulation run.",
            "**No** — strictly `<`",
            "recommended_iterations",
            "PCCM_ModelCheckCalculationState",
            "declared-production-correction discipline"):
        assert decision in text, f"the record does not settle: {decision}"
    # THE SUPERSEDED WORDING IS GONE.
    assert "results are valid but less precise" not in text
    # AND THE WINDOW IS NOT THE SENSITIVITY ONE.
    assert "row_window = 1000" not in text


def test_51_optional_outputs_are_information_not_warnings() -> None:
    text = _text()
    table = text[text.index("### 2.1 Severity assignment"):text.index("### 2.2")]
    for row, severity in (("No simulation published", "INFO"),
                          ("Annual `NOT PRODUCED`", "INFO"),
                          ("Sensitivity not produced for this run", "INFO"),
                          ("Calculation `NOT CALCULATED`", "WARNING"),
                          ("Calculation `STALE`", "WARNING"),
                          ("Simulation `STALE`", "WARNING"),
                          ("Annual `HISTORICAL` or `OTHER Px`", "WARNING")):
        line = [l for l in table.splitlines() if l.startswith(f"| {row}")]
        assert line, f"the taxonomy omits {row}"
        assert f"**{severity}**" in line[0], f"{row} is not {severity}: {line[0]}"


def test_52_the_actionable_count_rule_is_explicit_and_reconciles() -> None:
    text = _text()
    assert "A row is **actionable** if it is ERROR or WARNING." in text
    assert "INFO rows are never counted." in text
    assert "Overall Status reconciles exactly to the actionable rows." in text
    # AND THE SCENARIOS CARRY EXPLICIT COUNTS.
    matrix = text[text.index("## 11. Scenario matrix"):text.index("## 12. ")]
    for scenario in "ABCDEFGHIJ":
        assert f"| **{scenario}** |" in matrix, f"scenario {scenario} is missing"
    assert matrix.count("0 / 0") >= 2 and "1 / n" in matrix, (
        "the matrix does not state actionable counts")
