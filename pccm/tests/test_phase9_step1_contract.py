#!/usr/bin/env python3
"""The Phase-9 Step-1 contract says only things the source can be asked.

WHAT A CONTRACT RECORD IS FOR. It is the authority P9-2 will be built against.
Every claim it makes about an existing owner is therefore re-derived here, so a
plan built on it cannot rest on a fact that has quietly stopped being true.

AND IT IS A CONTRACT, NOT AN IMPLEMENTATION - which is a claim about the commit
that RECORDED it, not about every commit after it. As first written, two controls
scanned today's working tree for Phase-9 work and refused to find any. That made
a design record's honesty depend on nobody ever implementing it, and it went red
the hour P9-2 began, which is how it was found. Both now ask the STEP-1 COMMIT's
own tree, which is the tree the sentence is about, and both gained the other
half of the claim: what P9-2 actually built has to be what this record
authorised. That is stricter than "nothing exists yet" ever was.

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

# THE COMMIT THAT RECORDED THIS CONTRACT. Everything the record says about "the
# tree" is a statement about this tree.
STEP_1_COMMIT = "dd082c9"


def _git(*args: str) -> str:
    import subprocess

    return subprocess.run(["git", *args], cwd=PCCM_ROOT.parent, check=True,
                          stdout=subprocess.PIPE, text=True).stdout


def _at_step_1(path: str) -> str:
    import subprocess

    result = subprocess.run(["git", "show", f"{STEP_1_COMMIT}:{path}"],
                            cwd=PCCM_ROOT.parent, capture_output=True, text=True)
    return result.stdout if result.returncode == 0 else ""


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


def test_02_no_phase_9_implementation_existed_when_this_was_recorded() -> None:
    """THE CLAIM NOBODY WOULD OTHERWISE CHECK, asked of the commit that made it."""
    listing = _git("ls-tree", "-r", "--name-only", STEP_1_COMMIT, "pccm/").split()
    strays = [name for name in listing
              if "phase9" in name.rsplit("/", 1)[-1].lower()
              and name.split("/")[1] in ("builder", "bootstrap", "build", "spec", "src")]
    assert not strays, (
        f"the Step-1 commit {STEP_1_COMMIT} already carried an implementation while the "
        f"record said none did: {strays}")
    # AND THE TWO SPEC ADDITIONS IT PLANNED WERE NOT THERE YET.
    manifest = yaml.safe_load(_at_step_1("pccm/spec/workbook.yaml"))
    assert "phase9_shell" not in manifest, "phase9_shell already existed"
    inputs = yaml.safe_load(_at_step_1("pccm/spec/input_contract.yaml"))
    assert "recommended_iterations" not in inputs["inputs"]["monte_carlo_iterations"], (
        "the recommendation field already existed; §9.1 was out of date when written")


def test_02b_what_was_built_is_what_this_record_authorised() -> None:
    """THE OTHER HALF OF THE CLAIM. A plan is only worth recording if the thing
    built from it is the thing it described.

    WHAT IS ASSERTED, AND WHY NOT MORE. §13 is a PLAN - "files expected to
    change" - not a whitelist, so requiring that nothing else exist would forbid
    a control written for a file the plan did name, which is legitimate work.
    What is asserted is that every planned file exists once P9-2 has run, and
    that any additional Phase-9 file sits in a directory the plan already
    reaches. A Phase-9 file in a directory §13 never mentions is work the record
    did not describe, and that still fails."""
    expected = {line.strip() for line in
                _text().split("## 13.")[1].split("```text")[1].split("```")[0].splitlines()
                if line.strip()}
    planned = {line.split()[0] for line in expected if line and not line.startswith("#")}
    found = sorted(str(path.relative_to(PCCM_ROOT)) for area in
                   ("builder", "bootstrap", "spec", "src", "tests", "docs")
                   for path in (PCCM_ROOT / area).rglob("*phase9*")
                   if path.is_file() and "__pycache__" not in str(path))
    # THE RECORD AND ITS OWN CONTROLS ARE STEP 1's OWN ARTEFACTS. §13 lists what
    # P9-2 changes; it does not list the document making the list, or the file
    # you are reading.
    own = {"docs/phase9_step1_contract.md", "tests/test_phase9_step1_contract.py"}
    directories = {plan.rsplit("/", 1)[0] for plan in planned if "/" in plan}
    for name in found:
        if name in own:
            continue
        assert name.rsplit("/", 1)[0] in directories, (
            f"{name} is Phase-9 work in a directory the record's §13 never reaches")
    assert own <= set(found), sorted(found)

    # AND ONCE P9-2 HAS RUN, EVERY NEW FILE THE PLAN NAMED EXISTS. Until then
    # the plan is simply not finished, which is not a failure.
    if any(name.startswith("builder/") for name in found):
        missing = [plan for plan in planned
                   if plan.endswith((".py", ".ps1", ".json")) and "phase9" in plan
                   and not (PCCM_ROOT / plan).is_file()]
        assert not missing, f"§13 planned files P9-2 did not produce: {missing}"


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


def test_11_the_adapter_it_authorises_was_work_not_yet_done() -> None:
    """THE AUTHORISATION WAS FOR WORK NOT DONE, at the commit that authorised it."""
    for name in ("CalcReportDerivedStatus", "PCCM_ModelCheckCalculationState"):
        assert name in _text(), f"the record does not name {name}"
        assert name not in _at_step_1("pccm/src/vba/modCalcReport.bas"), (
            f"{name} already existed when the record authorised it")
        assert name not in _at_step_1("pccm/src/vba/modResultsState.bas"), (
            f"{name} already existed when the record authorised it")


def test_11b_the_authorised_adapter_was_built_where_it_was_authorised() -> None:
    """AND THE AUTHORISATION WAS HONOURED, OR NOT USED AT ALL. §8 names the owner
    module for each half; an adapter that landed somewhere else would be an
    authorisation quoted rather than followed."""
    calc, state = _module("modCalcReport.bas"), _module("modResultsState.bas")
    if "CalcReportDerivedStatus" not in calc and \
            "PCCM_ModelCheckCalculationState" not in state:
        pytest.skip("P9-2 has not built the adapter")
    assert "Public Function CalcReportDerivedStatus" in calc, (
        "the pure derivation was not exposed in modCalcReport, which §8 names")
    assert "Public Function PCCM_ModelCheckCalculationState" in state, (
        "the adapter is not in modResultsState, which §8 names as the owner")
    assert "modCalcReport.CalcReportDerivedStatus(" in state, (
        "the adapter does not delegate to the pure derivation §8 authorises")
    # AND PCCM_CalculationStatus IS UNTOUCHED, which §8 states in as many words.
    assert "WriteStatusBlock status" in _procedure("modCalcReport.bas",
                                                  "PCCM_CalculationStatus")


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
