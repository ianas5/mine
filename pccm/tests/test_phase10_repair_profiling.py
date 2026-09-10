#!/usr/bin/env python3
"""P10-2C: Repair Profiling.

WHAT THIS FILE PROVES, ON LINUX, WITHOUT EXCEL - AND IN TWO HALVES, BECAUSE THE
TWO CLAIMS ARE DIFFERENT KINDS OF CLAIM.

  WHAT A REPAIR PRODUCES is the accepted structural semantics, and this project
  already owns a specification of them: `structure_oracle` is pure Python over
  plain data and is what the Phase-4 Windows harness compares the real workbook
  against. Repair Profiling performs exactly two operations per grid -
  SetYearColumns then SyncRows - and the oracle has a function for each. So the
  behavioural controls below compose the ORACLE in that order and assert what
  comes out: a missing row restored, an orphan gone, order restored, a width
  changed, every attributable weight preserved by (permanent id, project year).

  WHAT THE COMMAND DECIDES is modRepair's own, and it is asserted against the
  source: which faults it repairs, which it refuses, what it captures before it
  mutates, what it restores, and what it never touches.

  THE JOIN BETWEEN THEM is test_02, which requires modRepair to call those two
  owners, in that order, with the plan's own start year and target width. That
  is what licenses the oracle composition to stand for the command.

NONE OF THIS RUNS EXCEL and none of it claims to. A clean Windows run proves the
runtime; this proves the design and the decisions.

Runs standalone or under pytest.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402
import yaml  # noqa: E402

from pccm_builder import load_structure_contract  # noqa: E402
from pccm_builder.structure_oracle import (  # noqa: E402
    orphan_rows,
    remap_profiling,
    removed_profiling_values,
    sync_profiling_values,
    sync_rows,
)
from pccm_builder.vba_source import VbaModule  # noqa: E402

SPEC = PCCM_ROOT / "spec"
SRC = PCCM_ROOT / "src" / "vba"
BUILD = PCCM_ROOT / "build"
GENERATED = BUILD / "vba"
REPAIR_BAS = SRC / "modRepair.bas"

ENDPOINT = "PCCM_RepairProfiling"
CAPTION = "Repair Profiling"
SHAPE = "btnPCCMRepairProfiling"

# THE TWO OWNERS THE REPAIR IS MADE OF, in the order it must call them.
OWNER_CALLS = ("modProfiling.SetYearColumns", "modProfiling.SyncRows")

# THE CONTRACT-DECLARED FAULT KEYS THIS COMMAND REPAIRS. Everything else the
# accepted checker can report is blocking, which is what makes a structural rule
# added later refuse by default instead of being silently ignored.
REPAIRABLE_FAULTS = (
    "CHK_PROFILING_COLUMN_COUNT",
    "CHK_PROFILING_YEAR_HEADERS",
    "CHK_COST_PROFILING_IDS_MATCH",
    "CHK_RISK_PROFILING_IDS_MATCH",
    "CHK_NO_DUPLICATE_IDS",
)

BLANK = None
INITIAL = 0.0

_CACHE: dict[str, object] = {}


def _module() -> VbaModule:
    if "module" not in _CACHE:
        _CACHE["module"] = VbaModule(name="modRepair", path=REPAIR_BAS,
                                     raw=REPAIR_BAS.read_bytes().decode("utf-8"))
    return _CACHE["module"]  # type: ignore[return-value]


def _code() -> str:
    return _module().code


def _procedure(name: str) -> str:
    source = _module().code_without_string_removal
    match = re.search(
        rf"^(?:Public|Private) (?:Sub|Function) {name}\b.*?^End (?:Sub|Function)$",
        source, re.M | re.S)
    assert match, f"modRepair.{name} not found"
    return match.group(0)


def _structure():
    if "structure" not in _CACHE:
        _CACHE["structure"] = load_structure_contract(SPEC / "structure_contract.yaml")
    return _CACHE["structure"]


def _generated_constants() -> dict[str, str]:
    if "constants" not in _CACHE:
        table: dict[str, str] = {}
        for path in sorted(GENERATED.glob("*.bas")):
            for line in path.read_text(encoding="utf-8").splitlines():
                match = re.match(r'^Public Const (\w+) As \w+ = (.+?)(?:\s{2,}\'.*)?$', line)
                if match:
                    table[match.group(1)] = match.group(2).strip().strip('"')
        _CACHE["constants"] = table
    return _CACHE["constants"]  # type: ignore[return-value]


# ===========================================================================
# THE ORACLE COMPOSITION - what a repair of ONE grid produces
# ===========================================================================
def repair_grid(register_ids: list[str], grid: dict[str, list[object]],
                target_years: int) -> dict[str, list[object]]:
    """SetYearColumns then SyncRows, in that order, over the accepted oracle.

    THE ORDER IS THE WHOLE OF IT. SyncRows preserves a weight by (permanent id,
    project-year INDEX), so the index set has to be the one the applied timeline
    calls for before ownership is re-established over it. Reversing these two
    would preserve weights into positions that are about to be reshaped.
    """
    reshaped = remap_profiling(grid, target_years, INITIAL)
    ordered, _added, _removed = sync_rows(register_ids, reshaped)
    return sync_profiling_values(reshaped, ordered, target_years, INITIAL)


def test_01_the_repair_is_two_owner_calls_and_nothing_else() -> None:
    """OWNERSHIP. modRepair holds no grid geometry: the repair is one call to
    the year-column owner and one to the row owner, both accepted, both
    byte-frozen, and both already driven this way by Apply / Update Timeline."""
    apply_body = _procedure("Apply")
    for call in OWNER_CALLS:
        assert call in apply_body, call
    assert apply_body.index(OWNER_CALLS[0]) < apply_body.index(OWNER_CALLS[1]), (
        "the rows are synchronised before the columns they are preserved into")
    # AND THE ARGUMENTS COME FROM THE PLAN, not from anything this module
    # decided for itself.
    assert "plan.StartYear, plan.TargetYears" in apply_body
    assert "plan.Kind" in apply_body
    # NOTHING ELSE IN THE MODULE WRITES A CELL. Reads are what an assessment is
    # made of; an ASSIGNMENT to a cell would be this command reaching past the
    # two owners above and writing a grid itself.
    assert not re.search(r"\.Value2?\s*=", _code()), _code()
    for banned in ("ListRows", "ListColumns.Add", "ClearContents", "Delete",
                   "ListColumns(", "DataBodyRange"):
        assert banned not in _code(), banned


def test_02_the_command_names_no_address_and_no_structural_rule() -> None:
    """OWNERSHIP, in the form that survives a later edit. modRepair may name the
    two grids' SHEETS - it reports about them - and the contract's fault keys,
    limits and names. It may not name a table, a column or a register field."""
    referenced = _module().referenced_upper_identifiers & set(_generated_constants())
    allowed = {
        "SH_COST_PROFILING", "SH_RISK_PROFILING",
        "NM_APPLIED_DURATION", "NM_APPLIED_START_YEAR",
        "LIMIT_MAX_YEAR_COLUMNS", "LIMIT_MIN_YEAR", "LIMIT_MAX_YEAR",
        "TOL_PROFILING_SUM_ABSOLUTE",
    } | set(REPAIRABLE_FAULTS)
    assert referenced == allowed, sorted(referenced ^ allowed)
    for banned in ("TBL_", "COL_", "ID_PREFIX", "GRID_", "PROFILE_INITIAL_VALUE"):
        assert banned not in _code(), banned


# ===========================================================================
# A. WHAT A REPAIR PRODUCES
# ===========================================================================
def test_03_a_correct_grid_is_left_exactly_as_it_is() -> None:
    """1. No-op on a structurally correct pair - and "no-op" means the values
    are identical, not merely that the command reports success."""
    grid = {"CL-001": [0.5, 0.5], "CL-002": [1.0, 0.0]}
    assert repair_grid(["CL-001", "CL-002"], grid, 2) == grid


def test_04_a_missing_row_is_restored_at_the_contract_default() -> None:
    """2 and 3. A driver with no profiling row gets one, and it starts at the
    contract's own initial value in every project year - the only place this
    command may write a weight, and only because the profiling owner does."""
    grid = {"CL-001": [0.5, 0.5]}
    repaired = repair_grid(["CL-001", "CL-002"], grid, 2)
    assert repaired == {"CL-001": [0.5, 0.5], "CL-002": [INITIAL, INITIAL]}
    # AND THE EXISTING ROW WAS NOT TOUCHED to make room for the new one.
    assert repaired["CL-001"] == grid["CL-001"]


def test_05_an_orphan_row_is_removed_and_takes_nothing_with_it() -> None:
    """4 and 5. A row whose id is no longer an identified driver is not a
    profile of anything; it goes, and no surviving row changes."""
    grid = {"CL-001": [0.4, 0.6], "CL-009": [1.0, 0.0]}
    repaired = repair_grid(["CL-001"], grid, 2)
    assert repaired == {"CL-001": [0.4, 0.6]}


def test_06_row_order_follows_the_register_and_transfers_nothing() -> None:
    """6. Reordering the register reorders the grid. Ownership is by permanent
    id, so a row that moves keeps its own weights and takes none from the row
    that used to sit where it lands."""
    grid = {"CL-002": [0.2, 0.8], "CL-001": [0.7, 0.3]}
    repaired = repair_grid(["CL-001", "CL-002"], grid, 2)
    assert list(repaired) == ["CL-001", "CL-002"]
    assert repaired["CL-001"] == [0.7, 0.3]
    assert repaired["CL-002"] == [0.2, 0.8]


def test_07_the_width_follows_the_applied_timeline() -> None:
    """7. Growth appends project years at the contract default; every existing
    position keeps its value, so a row that totalled 100% still does."""
    grid = {"CL-001": [0.5, 0.5]}
    grown = repair_grid(["CL-001"], grid, 4)
    assert grown["CL-001"] == [0.5, 0.5, INITIAL, INITIAL]
    assert sum(v for v in grown["CL-001"] if v is not None) == 1.0


def test_08_every_attributable_weight_survives_a_width_repair() -> None:
    """8 and 9. Weights belonging to SURVIVING project-year identities are kept
    exactly - including a blank, which is not the same as 0% and must not become
    one inside a structural operation."""
    grid = {"CL-001": [0.25, BLANK, 0.75], "CL-002": [1.0, 0.0, 0.0]}
    grown = repair_grid(["CL-001", "CL-002"], grid, 5)
    assert grown["CL-001"] == [0.25, BLANK, 0.75, INITIAL, INITIAL]
    assert grown["CL-002"] == [1.0, 0.0, 0.0, INITIAL, INITIAL]
    # AND A SHRINK KEEPS THE POSITIONS THAT SURVIVE. The command only reaches
    # this case when the trimmed tail carries no data - test_14 is that gate.
    shrunk = repair_grid(["CL-001"], {"CL-001": [0.6, 0.4, BLANK, 0.0]}, 2)
    assert shrunk["CL-001"] == [0.6, 0.4]


def test_09_a_repair_of_several_faults_at_once_is_still_value_preserving() -> None:
    """5. The realistic case: a missing row, an orphan, drifted order and a
    changed width, all in one grid. Every weight that can be attributed to a
    surviving driver and a surviving project year comes through unchanged."""
    grid = {
        "R-003": [0.1, 0.9],
        "R-001": [0.5, BLANK],
        "R-404": [1.0, 0.0],
    }
    repaired = repair_grid(["R-001", "R-002", "R-003"], grid, 3)
    assert list(repaired) == ["R-001", "R-002", "R-003"]
    assert repaired["R-001"] == [0.5, BLANK, INITIAL]
    assert repaired["R-002"] == [INITIAL, INITIAL, INITIAL]
    assert repaired["R-003"] == [0.1, 0.9, INITIAL]


def test_10_a_second_repair_changes_nothing() -> None:
    """17. Idempotence, proved by running the composition twice rather than by
    asserting the command reports a no-op."""
    grid = {"CL-002": [0.2, 0.8], "CL-009": [1.0, 0.0]}
    once = repair_grid(["CL-001", "CL-002"], grid, 3)
    twice = repair_grid(["CL-001", "CL-002"], once, 3)
    assert twice == once


def test_11_identical_duplicate_rows_collapse_and_lose_nothing() -> None:
    """10. Two rows for one id saying the SAME thing is one row said twice, and
    the oracle's row synchroniser is keyed by id so it cannot hold both. That is
    the only duplicate case the command allows - test_16 is the other one."""
    grid = {"CL-001": [0.5, 0.5]}
    assert repair_grid(["CL-001"], grid, 2) == grid
    # AND THE ORACLE REFUSES A REPEATED REGISTER ID OUTRIGHT, which is why
    # modRepair refuses it before ever reaching this composition.
    with pytest.raises(ValueError):
        sync_rows(["CL-001", "CL-001"], grid)


# ===========================================================================
# B. WHAT THE COMMAND REFUSES
# ===========================================================================
def _refusal(name: str) -> str:
    return _procedure(name)


def test_12_every_refusal_reports_through_the_accepted_result_type() -> None:
    """The refusal architecture: one OperationResult, through the accepted
    reporting owner, with no state machine of its own."""
    refused = _procedure("Refused")
    assert "modAppState.Failed(" in refused
    assert "Repair Profiling was refused. Nothing was changed." in _module().raw
    body = _procedure("RepairProfiling")
    assert body.count("Refused(detail)") == 4, body.count("Refused(detail)")
    # AND NO SECOND DIALOG, NO SECOND STATE, NO CONFIRMATION.
    assert "MsgBox" not in _code()
    assert "AskConfirm" not in _code(), (
        "Repair refuses what it cannot do safely; it has nothing to warn about")
    assert _procedure(ENDPOINT).count("modAppState.Announce") == 4


def test_13_an_unrecognised_structural_fault_blocks_the_command() -> None:
    """THE GATE THAT MAKES A LATER RULE SAFE. The accepted checker tags every
    fault with a contract key; this command repairs a named five and refuses the
    rest, so a structural rule added in a later phase is blocking by default."""
    body = _procedure("IsRepairableFault")
    named = set(re.findall(r"\bCHK_\w+", body))
    assert named == set(REPAIRABLE_FAULTS), sorted(named ^ set(REPAIRABLE_FAULTS))
    # EVERY KEY THE CHECKER CAN REPORT IS ACCOUNTED FOR: repaired, or blocking.
    declared = {name for name in _generated_constants() if name.startswith("CHK_")}
    assert set(REPAIRABLE_FAULTS) < declared, sorted(set(REPAIRABLE_FAULTS) - declared)
    blocking = declared - set(REPAIRABLE_FAULTS)
    assert "CHK_ID_PATTERN" in blocking
    assert "CHK_APPLIED_TRIPLE_CONSISTENT" in blocking
    assert "CHK_GRID_SHAPE" in blocking
    assert "CHK_NO_ORPHAN_STRUCTURAL_DATA" in blocking
    assert "CHK_COUNTER_INTEGRITY" in blocking
    gate = _procedure("OnlyRepairableFaults")
    assert "modStructuralCheck.ValidateStructure()" in gate
    # AND THE HEADER RULE IS TAKEN FROM THE CHECKER RATHER THAN ASKED AGAIN.
    # What a project-year header must read is decided in one place; a second
    # copy here would be this command becoming a structural authority.
    assert "CHK_PROFILING_YEAR_HEADERS" in gate
    assert "headerDrift = True" in gate
    assert "ListColumns" not in _code(), (
        "modRepair reads the grid's columns directly instead of asking an owner")


def test_14_unattributable_data_is_refused_by_the_accepted_owners() -> None:
    """14. Neither of the two "cannot be attributed" cases is decided here.

    A row with data and NO key is the accepted pre-mutation gate's question, and
    a trimmed tail carrying data is the profiling owner's own destructive
    assessment - the one Apply / Update Timeline warns from. Repair calls both
    and refuses on either.
    """
    body = _procedure("RepairProfiling")
    assert "modStructuralCheck.PreMutationCheck()" in body
    assert body.index("PreMutationCheck") < body.index("OnlyRepairableFaults")
    trim = _procedure("TrimIsEmpty")
    assert "modProfiling.CountDataBeyond" in trim
    assert "hits = 0" in trim
    # AND THE ORACLE AGREES ABOUT WHAT EITHER OF THEM FINDS.
    assert orphan_rows([(None, [0.5]), ("CL-001", [0.5])]) == [1]
    assert removed_profiling_values({"CL-001": [0.5, 0.5]}, 1) == [("CL-001", 2, 0.5)]
    assert removed_profiling_values({"CL-001": [1.0, BLANK, 0.0]}, 1) == []


def test_15_an_unreadable_identifier_or_weight_is_refused() -> None:
    """12. An id that is an error value keys nothing, and a weight that is an
    error value or text cannot be preserved as a weight. Both are refused by
    name, on both sides - the register's ids and the grid's."""
    register = _procedure("ReadRegisterIds")
    grid = _procedure("ReadGridIds")
    weights = _procedure("ReadRowWeights")
    for body in (register, grid):
        assert "modWorkbook.IsErrorText" in body
    assert "modWorkbook.IsErrorText" in weights
    assert "Not IsNumeric(cell.Value)" in weights
    assert "cannot be preserved" in weights


def test_16_a_duplicate_row_with_conflicting_weights_is_refused() -> None:
    """11. Two rows for one id saying DIFFERENT things are two answers to one
    question. Collapsing them would be this command choosing which of the user's
    two profiles they meant, and the id is named so they can choose instead."""
    body = _procedure("ReadGridIds")
    assert "SameWeights(held, weights)" in body
    assert "different weights" in body
    same = _procedure("SameWeights")
    # A BLANK IS NOT A ZERO, here as everywhere: the grid language distinguishes
    # "not entered" from "nothing allocated", and a comparison that collapsed
    # them would call two different profiles identical.
    assert "IsEmpty(held(index)) <> IsEmpty(found(index))" in same


def test_17_a_repeated_register_identifier_is_refused() -> None:
    """11. The other half of the duplicate question, and it is not repairable at
    all: two drivers claiming one id means a profiling row could belong to
    either, and nothing in the workbook says which."""
    body = _procedure("ReadRegisterIds")
    assert "more " in body and "than once" in body
    assert "ids.Exists(idText)" in body


def test_18_a_half_entered_profile_is_never_restructured_or_normalised() -> None:
    """13. THE SEMANTIC GATE, AND THE ONE THIS COMMAND MOST HAD TO GET RIGHT.

    Two totals are recognised: 100% is a profile, and 0% is the ABSENCE of one -
    which is what the contract seeds a new driver and a new project year with,
    so refusing on it would block every workbook with an unprofiled driver.
    Anything between them is half-entered, and it is refused only when the
    project-year columns are actually changing, because that is the only case
    where restructuring would move those weights between two different sets of
    positions.
    """
    gate = _procedure("RecognisedProfile")
    caller = _procedure("ReadGridIds")
    assert "plan.WidthDrift And registerIds.Exists(idText)" in caller
    # THE ARITHMETIC IS NOT THIS MODULE'S. Same summation primitive, same
    # subtraction and same contract tolerance the Model Check uses to ask the
    # same question.
    assert "modCalcFactors.SafeSignedSum" in gate
    assert "modCalcFactors.SafeSubtract" in _procedure("WithinTolerance")
    assert "TOL_PROFILING_SUM_ABSOLUTE" in _procedure("WithinTolerance")
    assert "never" in gate and "chooses weights" in gate
    # AND NOTHING NORMALISES. A repair that scaled a row to 100% would be
    # inventing an allocation.
    for banned in ("/ total", "/total", "Normalise", "Normalize", "Rescale"):
        assert banned not in _code(), banned


def test_19_the_two_recognised_totals_are_the_ones_the_model_check_uses() -> None:
    """13. The target is 100%, and it is the SAME 100% modCalcCheck compares a
    resolved profile against. The two literals are compared directly, so they
    cannot drift into two different definitions of what a profile is."""
    checker = (SRC / "modCalcCheck.bas").read_text(encoding="utf-8")
    theirs = re.search(r"Private Const PROFILE_SUM_TARGET As Double = (\S+)", checker)
    ours = re.search(r"Private Const REPAIR_PROFILE_SUM_TARGET As Double = (\S+)",
                     _module().raw)
    assert theirs and ours, (theirs, ours)
    assert theirs.group(1) == ours.group(1), (theirs.group(1), ours.group(1))
    # AND THE EMPTY TOTAL IS THE CONTRACT'S OWN INITIAL VALUE - what
    # SetYearColumns seeds a new project year with - so "no profile yet" means
    # the same thing to this command as it does to the profiling owner. The
    # trailing "#" is VBA's Double type character and is not part of the number.
    empty = re.search(r"Private Const REPAIR_PROFILE_SUM_EMPTY As Double = (\S+)",
                      _module().raw)
    assert empty, _module().raw
    assert float(empty.group(1).rstrip("#")) == float(
        _generated_constants()["PROFILE_INITIAL_VALUE"])


# ===========================================================================
# C. THE TRANSACTION
# ===========================================================================
def test_20_both_grids_are_captured_before_either_is_touched() -> None:
    """15. "Cost repaired and Risk broken" is the outcome the transaction exists
    to make impossible, so both snapshots are taken before the first Apply and
    both assessments are made before either snapshot."""
    body = _procedure("RepairProfiling")
    first_assess = body.index("Assess(modProfiling.CostKind()")
    second_assess = body.index("Assess(modProfiling.RiskKind()")
    first_capture = body.index("costBefore = modWorkbook.SnapshotTable")
    second_capture = body.index("riskBefore = modWorkbook.SnapshotTable")
    first_apply = body.index("Apply cost")
    assert first_assess < second_assess < first_capture < second_capture < first_apply
    assert "captured = True" in body
    assert body.index("captured = True") < first_apply


def test_21_a_failure_anywhere_restores_both_grids() -> None:
    """16. The snapshot pair is the whole of what this command can change, so
    restoring it is the whole rollback - and the failure says which of the three
    outcomes happened rather than claiming the best one."""
    body = _procedure("Rollback")
    assert body.count("modWorkbook.RestoreTable") == 2
    assert "If Not captured Then" in body
    assert "Nothing had been modified" in body
    assert "restored to the state they were in before the" in body
    assert "could not be fully restored" in body
    # THE FAILPOINTS SIT BETWEEN THE TWO GRIDS AND AFTER BOTH, which is the
    # boundary a harness has to be able to fail at.
    stages = re.findall(r"modAppState\.FailPointCheck (FAILPOINT_REPAIR_\w+)",
                        _procedure("RepairProfiling"))
    assert stages == ["FAILPOINT_REPAIR_COST", "FAILPOINT_REPAIR_RISK"], stages
    for stage in stages:
        assert f"Public Const {stage} As String" in _module().raw, stage
    # AND NO SUPPRESSION ANYWHERE.
    assert "On Error Resume Next" not in _code()
    handlers = set(re.findall(r"On Error GoTo (\w+)", _code())) - {"0"}
    assert handlers == {"InvocationFailed", "NormalCleanupFailed", "CleanupFailed",
                        "RepairFailed", "RestoreFailed"}, sorted(handlers)


def test_22_the_accepted_validator_gates_the_result() -> None:
    """15. A repair that did not produce a structurally valid workbook is a
    failure and rolls back, judged by the accepted checker rather than by this
    command's own opinion of its work."""
    body = _procedure("RepairProfiling")
    assert "modStructuralCheck.ValidateStructure()" in body
    assert body.index("Apply risk") < body.index(
        "detail = modStructuralCheck.ValidateStructure()")
    assert "Err.Raise" in body, "a failed revalidation must reach the rollback"


# ===========================================================================
# D. STATE, PROTECTION AND THE BUTTON
# ===========================================================================
def test_23_no_state_word_is_written_and_no_fingerprint_is_touched() -> None:
    """18 and 19. Repair Profiling forces nothing. It writes no state word, no
    fingerprint, no publication and no attempt record; the existing owners
    derive what they always derived, from inputs this command preserves.

    A structural-only repair that preserved the same weights therefore cannot
    create artificial staleness: the calculation fingerprint covers the resolved
    inputs, and the resolved inputs are the same weights against the same
    project years.
    """
    literals = re.findall(r'"([^"]*)"', _module().raw)
    for word in ("NOT CALCULATED", "CURRENT", "STALE", "INVALID", "NOT PRODUCED",
                 "PUBLISHED", "No timeline applied", "STRUCTURE CHANGE PENDING",
                 "Timeline current"):
        assert word not in literals, f"modRepair writes the state word {word}"
    for banned in ("Fingerprint", "modCalcReport", "modCalcFingerprint",
                   "modSimReport", "modSimFingerprint", "modResultsState",
                   "STATE_CURRENT", "STATE_PENDING", "STATE_NOT_APPLIED",
                   "NM_STRUCTURAL_STATE", "CALC_", "SIM_"):
        assert banned not in _code(), banned
    # AND IT WRITES NOTHING OUTSIDE THE TWO PROFILING GRIDS. modWorkbook is
    # reached only for reads and for the snapshot pair.
    calls = set(re.findall(r"modWorkbook\.(\w+)", _code()))
    assert calls <= {"SnapshotTable", "RestoreTable", "BodyRowCount", "CellIn",
                     "TextOf", "IsErrorText", "IsEmptyCell", "IsWholeInRange",
                     "SafeLong", "ReadLongInRange"}, sorted(calls)


def test_24_the_command_never_unprotects_anything() -> None:
    """20 and 21. Protection is UserInterfaceOnly, so every write the owners make
    already succeeds. An Unprotect here would leave the workbook open if the
    command then failed."""
    for banned in ("Unprotect", ".Protect", "modProtection", "Password"):
        assert banned not in _code(), banned
    # AND IT CHANGES NO SHEET, NAME, SHAPE OR TABLE STRUCTURE OF ITS OWN.
    for banned in ("Shapes", "Visible", "Names.Add", "ListObjects.Add"):
        assert banned not in _code(), banned


def test_25_the_button_is_declared_bound_and_last_in_the_block() -> None:
    """22, 23 and 24. The sixth and last command the contract authorises."""
    structure = _structure()
    buttons = {b.shape_name: b for b in structure.buttons}
    assert SHAPE in buttons
    button = buttons[SHAPE]
    assert (button.caption, button.entry_point, button.sheet) == (
        CAPTION, ENDPOINT, "Setup")
    assert ENDPOINT in structure.entry_points
    assert f"Public Sub {ENDPOINT}()" in _module().raw
    assert ENDPOINT not in structure.api_procedures
    block = structure.commands["block"]
    column, first = str(block["button_column"]), int(block["first_button_row"])
    pitch = int(block["button_row_pitch"])
    commands = [b for b in structure.buttons if b.sheet == block["sheet"]]
    assert [b.anchor_cell for b in commands] == [
        f"{column}{first + pitch * i}" for i in range(len(commands))]
    assert [b.entry_point for b in commands] == [
        "PCCM_ApplyTimeline", "PCCM_Calculate", "PCCM_RunSimulation",
        "PCCM_RunSensitivity", "PCCM_RunAnnualStochastic", "PCCM_ResetResults",
        ENDPOINT]
    # AND EVERY STRUCTURAL ADD/DELETE BUTTON IS UNTOUCHED.
    registers = [b for b in structure.buttons if b.sheet != block["sheet"]]
    assert [(b.sheet, b.caption, b.anchor_cell) for b in registers] == [
        ("Cost Lines", "Add Cost Line", "N6"),
        ("Cost Lines", "Delete Cost Line", "N9"),
        ("Risk Register", "Add Risk", "O6"),
        ("Risk Register", "Delete Risk", "O9")]


def test_26_the_success_wording_says_what_happened_and_what_was_kept() -> None:
    """USER FEEDBACK. A no-op says the grids were already correct; a repair says
    what was rebuilt and that attributable weights were preserved. Neither
    sentence exists anywhere else."""
    raw = _module().raw
    assert "are already " in raw and "structurally correct" in raw
    summary = _procedure("Summary")
    assert "preserved" in summary and "permanent identifier and project year" in summary
    describe = _procedure("Describe")
    for finding in ("missing row(s) restored", "orphan row(s) removed",
                    "duplicate row(s) collapsed", "row order restored",
                    "project-year columns "):
        assert finding in describe, finding
    for path in sorted(SRC.glob("*.bas")):
        if path.name != "modRepair.bas":
            assert "structurally correct" not in path.read_text(encoding="utf-8"), path.name


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
