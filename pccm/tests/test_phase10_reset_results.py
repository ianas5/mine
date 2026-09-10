#!/usr/bin/env python3
"""P10-2B: Reset Results.

WHAT THIS FILE PROVES, ON LINUX, WITHOUT EXCEL.

Reset Results clears every published result and preserves every input and every
identity value. Neither half can be demonstrated by running it here, so each is
proved in the one form a static reading can make exact:

  THE CLEAR SCOPE IS THE PUBLICATION SCOPE, because the clear is written in the
  SAME TERMS the publication is - the same table constants, the same range
  functions, the same ClearRecords - inside the same module. It is not a second
  description of where a result lives that happens to agree today.

  THE PRESERVED SCOPE IS SAFE BY ABSENCE. No input, counter or nonce constant is
  named anywhere in the new code. There is no preservation routine to get wrong;
  there is nothing that could reach an input in the first place.

  NO STATE WORD IS WRITTEN. The words a user reads after a reset - NOT
  CALCULATED, NOT PRODUCED, and the sentence Sensitivity prints - appear nowhere
  in what this command writes. They are derived by the owners that own them, from
  the absence of a publication.

AND THE ORCHESTRATOR HOLDS NO GEOMETRY. modReset names four owners and the order
it asks them in. It names no bank, no column, no row, no table and no address,
so it cannot become a second opinion about where anything lives.

Runs standalone or under pytest.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

from pccm_builder import load_structure_contract  # noqa: E402
from pccm_builder.vba_source import VbaModule  # noqa: E402
from vba_reset_plumbing import reset_addition_of  # noqa: E402

SRC = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"
BUILD = PCCM_ROOT / "build"
GENERATED = BUILD / "vba"
RESET_BAS = SRC / "modReset.bas"
PROJECTION = BUILD / "phase10_reset_inspection.json"

# (owner module, the pair it gained). The prefix is the owner's own, so a reader
# of modReset can see which store each call reaches without opening it.
OWNERS = (
    ("modCalcReport", "CalcReportClearPublication", "CalcReportRestorePublication"),
    ("modSimReport", "SimReportClearPublication", "SimReportRestorePublication"),
    ("modSimAnnualStore", "SimAnnualStoreClearPublication",
     "SimAnnualStoreRestorePublication"),
    ("modSimPostReport", "SimPostReportClearPublication",
     "SimPostReportRestorePublication"),
)

ENDPOINT = "PCCM_ResetResults"
CAPTION = "Reset Results"
SHAPE = "btnPCCMResetResults"

# The two orthogonal state vocabularies, plus the annual and sensitivity words.
# NOTHING RESET WRITES MAY BE ONE OF THEM: they are derived readings, and a
# command that wrote one would be a second opinion about state.
STATE_WORDS = ("NOT CALCULATED", "CURRENT", "STALE", "INVALID", "NOT PRODUCED",
               "HISTORICAL", "OTHER Px", "PUBLISHED")

# The identity that must survive a reset, by the constant that names it.
IDENTITY_CONSTANTS = (
    "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE",
    "SIM_IDENTITY_ROW_LAST_RUN_ID",
    "SIM_PENDING_AUTO_NONCE_CELL",
    "SIM_FINAL_COMMIT_RANGE",
)

_CACHE: dict[str, object] = {}


def _module(name: str) -> VbaModule:
    key = f"module:{name}"
    if key not in _CACHE:
        path = SRC / f"{name}.bas"
        _CACHE[key] = VbaModule(name=name, path=path,
                                raw=path.read_bytes().decode("utf-8"))
    return _CACHE[key]  # type: ignore[return-value]


def _reset() -> VbaModule:
    return _module("modReset")


def _addition(name: str) -> str:
    """The declared P10-2B block of an owner, comments and strings removed."""
    key = f"addition:{name}"
    if key not in _CACHE:
        raw = reset_addition_of(_module(name).raw)
        assert raw, f"{name} carries no declared P10-2B block"
        _CACHE[key] = VbaModule(name=name, path=_module(name).path, raw=raw).code
    return _CACHE[key]  # type: ignore[return-value]


def _addition_raw(name: str) -> str:
    return reset_addition_of(_module(name).raw)


def _procedure(module: str, name: str) -> str:
    source = _module(module).code_without_string_removal
    match = re.search(
        rf"^(?:Public|Private) (?:Sub|Function) {name}\b.*?^End (?:Sub|Function)$",
        source, re.M | re.S)
    assert match, f"{module}.{name} not found"
    return match.group(0)


def _structure():
    if "structure" not in _CACHE:
        _CACHE["structure"] = load_structure_contract(SPEC / "structure_contract.yaml")
    return _CACHE["structure"]


def _projection() -> dict:
    if "projection" not in _CACHE:
        _CACHE["projection"] = json.loads(PROJECTION.read_text(encoding="utf-8"))
    return _CACHE["projection"]  # type: ignore[return-value]


def _generated_constants() -> dict[str, str]:
    if "constants" not in _CACHE:
        table: dict[str, str] = {}
        for path in sorted(GENERATED.glob("*.bas")):
            for line in path.read_text(encoding="utf-8").splitlines():
                match = re.match(r'^Public Const (\w+) As \w+ = (.+?)(?:\s{2,}\'.*)?$',
                                 line)
                if match:
                    table[match.group(1)] = match.group(2).strip().strip('"')
        _CACHE["constants"] = table
    return _CACHE["constants"]  # type: ignore[return-value]


# ===========================================================================
# A. THE OWNERS, AND THE ONE THING EACH OF THEM GAINED
# ===========================================================================
@pytest.mark.parametrize("owner,clear,restore", OWNERS)
def test_01_each_publication_owner_gained_exactly_a_clear_and_a_restore(
        owner: str, clear: str, restore: str) -> None:
    """§4. The audit's conclusion, asserted: no owner exposed a safe clear, so
    each gained the smallest one that does the job - and nothing else."""
    module = _module(owner)
    added = re.findall(r"^Public (?:Function|Sub) (\w+)", _addition_raw(owner), re.M)
    assert added == [clear, restore], added
    for name in (clear, restore):
        assert name in module.public_procedures, name
        assert not name.startswith("PCCM_"), (
            f"{name} is not an endpoint and must not look like one")
    # THE PAIR IS A FUNCTION PAIR WITH THE SAME SHAPE, so the orchestrator can
    # treat all four owners identically and needs no per-owner special case.
    for name in (clear, restore):
        assert re.search(
            rf"Public Function {name}\(ByRef undo As Variant, _\s*\n\s*"
            rf"ByRef detail As String\) As Boolean", _addition_raw(owner)), name


@pytest.mark.parametrize("owner,clear,restore", OWNERS)
def test_02_the_owner_captures_before_it_mutates(owner: str, clear: str,
                                                 restore: str) -> None:
    """§5. Only the backup necessary to restore what this command clears, taken
    under the same handler as the clear so a capture that raises has changed
    nothing - and HANDED OVER before the first mutation, so a clear that raises
    half way finds the carrier filled rather than empty."""
    body = _procedure(owner, clear)
    assert "On Error GoTo ClearFailed" in body
    assert "captured(index) = CapturedBlock(" in body
    handover = body.index("undo = captured")
    assert body.index("captured(index) = CapturedBlock(") < handover
    mutations = [body.index(token) for token in (".ClearContents", ".Value2 = ")
                 if token in body]
    assert mutations, f"{clear} clears nothing"
    assert handover < min(mutations), f"{clear} mutates before it hands over the undo"


@pytest.mark.parametrize("owner,clear,restore", OWNERS)
def test_03_the_restore_is_a_no_op_on_a_carrier_that_was_never_filled(
        owner: str, clear: str, restore: str) -> None:
    """§5. The orchestrator asks EVERY owner on the failure path rather than
    remembering which ones it reached, so "nothing was captured" has to be an
    answer rather than an error."""
    body = _procedure(owner, restore)
    assert "If IsEmpty(undo) Then" in body, restore
    guard = body.index("If IsEmpty(undo) Then")
    assert guard < body.index("On Error GoTo RestoreFailed")
    assert re.search(rf"{restore} = True\s*\n\s*Exit Function", body)


@pytest.mark.parametrize("owner,clear,restore", OWNERS)
def test_04_neither_procedure_suppresses_anything(owner: str, clear: str,
                                                  restore: str) -> None:
    """§5. A suppressed error in a destructive command is a silent half-reset."""
    addition = _addition(owner)
    assert "On Error Resume Next" not in addition
    handlers = set(re.findall(r"On Error GoTo (\w+)", addition)) - {"0"}
    assert handlers == {"ClearFailed", "RestoreFailed"}, sorted(handlers)


def test_05_the_calculation_clear_covers_what_a_calculation_writes() -> None:
    """§1. The five analytical tables, the totals block and the state block -
    named with the SAME constants WriteAnalytical and WriteSuccessCommit use."""
    forward = (_procedure("modCalcReport", "WriteAnalytical")
               + _procedure("modCalcReport", "WriteSuccessCommit"))
    written = {name for name in re.findall(r"\bTBL_CALC_\w+", forward)}
    written |= {name for name in re.findall(r"\bCALC_\w+_VALUE_RANGE\b", forward)}
    blocks = _procedure("modCalcReport", "PublicationBlocks")
    for constant in written:
        assert constant in blocks, f"the reset does not clear {constant}"
    assert len(re.findall(r"BodyAddress\(TBL_CALC_\w+\)", blocks)) == 5, blocks
    assert "CALC_TOTALS_VALUE_RANGE" in blocks and "CALC_STATE_VALUE_RANGE" in blocks
    # AND THE TABLE BODIES ARE ASKED FOR BY ADDRESS, not resized. The forward
    # path owns the row count; a reset that changed the shape as well as the
    # contents would have a shape to put back.
    assert "ResizeBody" not in _addition("modCalcReport")


def test_06_the_simulation_clear_uses_the_publications_own_range_builders() -> None:
    """§1 and §4. Not a second description of the banks: the same four range
    functions the candidate publication and the iteration writer already use."""
    blocks = _procedure("modSimReport", "PublicationBlocks")
    forward = (_procedure("modSimReport", "PublishCandidate")
               + _procedure("modSimReport", "WriteIterationBank"))
    for builder in ("SnapshotRange", "SummaryRange", "ContingencyRange",
                    "IterationRange"):
        assert builder in blocks, builder
        assert builder in forward, builder


def test_07_both_simulation_banks_are_cleared_and_neither_is_the_active_one() -> None:
    """§1 and §14. Clearing only the bank the selector names would leave a whole
    published distribution on the sheet, reachable the moment the selector moved.
    """
    blocks = _procedure("modSimReport", "PublicationBlocks")
    for builder in ("SnapshotRange", "SummaryRange", "ContingencyRange",
                    "IterationRange"):
        banks = re.findall(rf"{builder}\(\s*(SIM_BANK_[AB])", blocks)
        assert set(banks) == {"SIM_BANK_A", "SIM_BANK_B"}, (builder, banks)
    # AND THE ACTIVE-BANK SELECTOR IS NEVER CONSULTED to decide what to clear.
    clear = _procedure("modSimReport", "SimReportClearPublication")
    assert "ReadActiveBank" not in clear + blocks
    assert "ActiveSnapshotText" not in clear + blocks


def test_08_the_annual_clear_covers_both_banks_and_unpublishes_first() -> None:
    """§1. The marker goes blank before anything is removed and comes back last,
    exactly as the publication does it, so an interruption at either end leaves
    the block unreadable rather than half-true."""
    blocks = _procedure("modSimAnnualStore", "PublicationBlocks")
    clear = _procedure("modSimAnnualStore", "SimAnnualStoreClearPublication")
    for bank in ("SIM_BANK_A", "SIM_BANK_B"):
        assert f"StampAddress({bank})" in blocks, bank
        assert f"BlockAddress({bank})" in blocks, bank
        assert f"StampCell({bank}, SIM_ANNUAL_STAMP_ROW_PUBLISHED)" in clear, bank
    assert clear.index("SIM_ANNUAL_STAMP_ROW_PUBLISHED") < clear.index("ClearContents")
    # THE STAMPS ARE FIRST IN THE LIST AND THE RESTORE WALKS IT BACKWARDS, which
    # is what puts the PUBLISHED marker back last.
    assert blocks.index("StampAddress") < blocks.index("BlockAddress")
    restore = _procedure("modSimAnnualStore", "SimAnnualStoreRestorePublication")
    assert "For index = UBound(undo) To LBound(undo) Step -1" in restore


def test_09_the_sensitivity_clear_goes_through_the_publications_own_clear() -> None:
    """§4. The record height a reset covers is the height ClearRecords covers,
    written in the same terms, so a shrunken model cannot leave surplus rows
    behind after a reset either."""
    blocks = _procedure("modSimPostReport", "PublicationBlocks")
    clear = _procedure("modSimPostReport", "SimPostReportClearPublication")
    records = _procedure("modSimPostReport", "RecordAddress")
    accepted = _procedure("modSimPostReport", "ClearRecords")
    for bank in ("SIM_BANK_A", "SIM_BANK_B"):
        assert f"StampAddress({bank})" in blocks, bank
        assert f"RecordAddress({bank})" in blocks, bank
        assert f"StampCell({bank}, SIM_SENSITIVITY_STAMP_ROW_PUBLISHED)" in clear, bank
    for term in ("SIM_SENSITIVITY_FIRST_ROW", "SIM_MAX_ITERATIONS"):
        assert term in records and term in accepted, term
    assert "SensitivityFirstColumn" in records and "SensitivityLastColumn" in records
    restore = _procedure("modSimPostReport", "SimPostReportRestorePublication")
    assert "For index = UBound(undo) To LBound(undo) Step -1" in restore


def test_10_nothing_is_left_hidden_rather_than_cleared() -> None:
    """§1: "Do not leave a stale publication merely hidden from presentation."
    Each store's RECORDS are in the clear scope, not only the marker that
    presents them - and each is cleared to the CONTRACT ceiling rather than to
    the height the last successful run happened to reach."""
    assert "IterationRange(SIM_BANK_A, 0, SIM_MAX_ITERATIONS)" in _procedure(
        "modSimReport", "PublicationBlocks")
    assert "LIMIT_MAX_YEAR_COLUMNS" in _procedure("modSimAnnualStore", "BlockAddress")
    assert "SIM_MAX_ITERATIONS" in _procedure("modSimPostReport", "RecordAddress")
    assert "BodyAddress" in _procedure("modCalcReport", "PublicationBlocks")


# ===========================================================================
# C. WHAT MUST SURVIVE, AND WHY IT CANNOT BE REACHED
# ===========================================================================
def test_11_the_simulation_reset_range_stops_above_the_counters() -> None:
    """§2 and §14. THE SINGLE MOST IMPORTANT ADDRESS IN THIS BATCH.

    The final commit writes D22:D30 in one assignment and D22 is the run-id
    counter; the AUTO nonce counter is D21 and its pending marker is F21. The
    reset range starts one row BELOW the commit range, and this resolves both
    from the generated contract rather than trusting the constant names to mean
    what they say.
    """
    constants = _generated_constants()
    first = int(constants["SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT"])
    last = int(constants["SIM_IDENTITY_ROW_ACTIVE_BANK"])
    nonce_row = int(constants["SIM_IDENTITY_ROW_NEXT_AUTO_NONCE"])
    run_id_row = int(constants["SIM_IDENTITY_ROW_LAST_RUN_ID"])
    assert nonce_row < first and run_id_row < first, (nonce_row, run_id_row, first)
    assert first <= last
    commit = constants["SIM_FINAL_COMMIT_RANGE"]
    column = constants["SIM_SHARED_VALUE_COLUMN"]
    assert commit == f"{column}{run_id_row}:{column}{last}", commit
    # AND THE SOURCE BUILDS ITS RANGE FROM THOSE TWO ROWS AND NO OTHER. The
    # contract being right is not the claim; the claim is that the reset range
    # starts at the attempt row.
    record = _procedure("modSimReport", "PublicationRecordRange")
    assert "SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT" in record
    assert "SIM_IDENTITY_ROW_ACTIVE_BANK" in record
    for banned in IDENTITY_CONSTANTS:
        assert banned not in record, banned
    # AND THE PROJECTION SAYS THE SAME THING, from the same contract.
    simulation = _projection()["publications"]["simulation"]
    assert simulation["cleared"]["attempt_and_selector"] == \
        f"{column}{first}:{column}{last}"
    assert simulation["preserved"] == {
        "next_auto_nonce": f"{column}{nonce_row}",
        "last_run_id": f"{column}{run_id_row}",
        "pending_auto_nonce": constants["SIM_PENDING_AUTO_NONCE_CELL"],
    }


@pytest.mark.parametrize("owner,clear,restore", OWNERS)
def test_12_no_owner_names_an_identity_it_must_preserve(owner: str, clear: str,
                                                        restore: str) -> None:
    """§2. Safe by absence: the constants that name a counter, a nonce or the
    commit block appear nowhere in what this command executes."""
    addition = _addition(owner)
    for constant in IDENTITY_CONSTANTS:
        assert constant not in addition, f"{owner} names {constant}"


def test_13_nothing_in_this_batch_can_reach_an_input() -> None:
    """§2. Safe by absence, and stated as an EXACT allow-list rather than a list
    of things to avoid.

    Every generated constant the new code may name is enumerated below, per
    owner. A register, a grid, an input cell, a counter or a nonce constant
    cannot appear because ANY constant outside its owner's list fails here - so
    this does not depend on somebody having thought of the particular input a
    future edit might reach for.
    """
    allowed = {
        "modCalcReport": {
            "CALC_SHEET", "CALC_STATE_VALUE_RANGE", "CALC_TOTALS_VALUE_RANGE",
            "CALC_STATE_ROW_LAST_ATTEMPT_RESULT", "CALC_ATTEMPT_NONE",
            "TBL_CALC_YEARS", "TBL_CALC_INFLATION_FACTORS", "TBL_CALC_FX",
            "TBL_CALC_DRIVERS", "TBL_CALC_ANNUAL",
        },
        "modSimReport": {
            "SIM_BANK_A", "SIM_BANK_B", "SIM_MAX_ITERATIONS",
            "SIM_SHARED_VALUE_COLUMN",
            "SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT", "SIM_IDENTITY_ROW_ACTIVE_BANK",
            "SIM_ATTEMPT_NONE",
        },
        "modSimAnnualStore": {
            "SIM_BANK_A", "SIM_BANK_B", "SIM_ANNUAL_FIRST_ROW",
            "SIM_ANNUAL_STAMP_ROW_RUN_ID", "SIM_ANNUAL_STAMP_ROW_PUBLISHED",
            "LIMIT_MAX_YEAR_COLUMNS",
        },
        "modSimPostReport": {
            "SIM_BANK_A", "SIM_BANK_B", "SIM_SENSITIVITY_FIRST_ROW",
            "SIM_SENSITIVITY_STAMP_ROW_RUN_ID", "SIM_SENSITIVITY_STAMP_ROW_PUBLISHED",
            "SIM_MAX_ITERATIONS",
        },
    }
    generated = set(_generated_constants())
    for owner, _clear, _restore in OWNERS:
        module = VbaModule(name=owner, path=_module(owner).path,
                           raw=_addition_raw(owner))
        named = module.referenced_upper_identifiers & generated
        assert named == allowed[owner], (owner, sorted(named ^ allowed[owner]))
    assert not (_reset().referenced_upper_identifiers & generated)

    # AND NO OWNER OF AN INPUT IS CALLED FROM ANY OF IT.
    input_owners = ("modDrivers", "modProfiling", "modInflation", "modTimeline",
                    "modSimNonce", "modStructuralCheck")
    for body in [_reset().code] + [_addition(owner) for owner, _c, _r in OWNERS]:
        for name in input_owners:
            assert name not in body, name
    structure = _structure()
    names = {counter.defined_name for counter in structure.counters}
    names |= {field.defined_name for field in structure.applied}
    names |= {field.defined_name for field in structure.derived}
    for body in [_reset().code] + [_addition(owner) for owner, _c, _r in OWNERS]:
        for name in names:
            assert name not in body, name


def test_14_the_preservation_manifest_is_derived_and_carries_no_volatile_cell() -> None:
    """§11. Built from the contracts, and deliberately excluding every formula:
    a run that compared a recomputed cell before and after a reset would fail on
    a workbook that was behaving correctly."""
    structure = _structure()
    preserved = _projection()["preserved"]
    assert preserved["applied_timeline"]["cells"] == [f.cell for f in structure.applied]
    assert preserved["permanent_id_counters"]["cells"] == [
        counter["cell"] for counter in structure.identity_block["counters"]]
    volatile = {field.cell for field in structure.derived}
    volatile.add(structure.structural_state.cell)
    everything = set(preserved["applied_timeline"]["cells"])
    everything |= set(preserved["permanent_id_counters"]["cells"])
    for block in preserved["editable_inputs"]:
        everything |= set(block["cells"])
    assert not (everything & volatile), sorted(everything & volatile)
    # AND IT IS NOT EMPTY. A manifest with nothing in it would pass every
    # comparison a Windows run could make.
    assert len(everything) > 100, len(everything)


def test_15_the_publication_manifest_names_every_store_a_run_must_inspect() -> None:
    """§12. Inspect the authoritative stores, never the presentation."""
    cleared = _projection()["publications"]["simulation"]["cleared"]
    assert set(cleared) == {
        "bank_snapshots", "attempt_and_selector", "summary", "contingency",
        "iteration_banks", "annual_stamps", "annual_blocks",
        "sensitivity_stamps", "sensitivity_records",
    }, sorted(cleared)
    for group in ("bank_snapshots", "summary", "contingency", "iteration_banks",
                  "annual_stamps", "annual_blocks", "sensitivity_stamps",
                  "sensitivity_records"):
        assert set(cleared[group]) == {"A", "B"}, group
    calculation = _projection()["publications"]["calculation"]["cleared"]
    assert len(calculation["tables"]) == 5
    assert calculation["state"] and calculation["totals"]
    # NO SHEET OF THE PRESENTATION IS IN IT.
    sheets = {_projection()["publications"]["calculation"]["sheet"],
              _projection()["publications"]["simulation"]["sheet"]}
    assert sheets == {"_Calc", "_SimData"}, sheets


# ===========================================================================
# D. NO STATE WORD, NO COUNTER, NO SECOND OPINION
# ===========================================================================
def test_16_no_state_word_is_written_anywhere_in_this_batch() -> None:
    """§3. "Do not write state words merely to create the desired display."

    The words a user reads after a reset are DERIVED by the owners that own them.
    The two words this batch does write are the attempt-axis initials the
    contracts declare a built workbook carries, and they are checked against the
    contracts in test_17 rather than assumed.
    """
    bodies = {"modReset": _reset().raw}
    bodies.update({owner: _addition_raw(owner) for owner, _c, _r in OWNERS})
    for name, raw in bodies.items():
        literals = re.findall(r'"([^"]*)"', raw)
        for word in STATE_WORDS:
            assert word not in literals, f"{name} writes the state word {word}"


def test_17_the_only_written_value_is_the_contracts_own_as_built_initial() -> None:
    """§3 and §8. Restoring a field to the value the contract says a built
    workbook carries is not choosing a state - and it is what makes a second
    reset a fixpoint."""
    constants = _generated_constants()
    calc_clear = _procedure("modCalcReport", "CalcReportClearPublication")
    sim_clear = _procedure("modSimReport", "SimReportClearPublication")
    assert "= CALC_ATTEMPT_NONE" in calc_clear
    assert "= SIM_ATTEMPT_NONE" in sim_clear
    projected = _projection()["publications"]["calculation"]["attempt_result_initial"]
    assert constants["CALC_ATTEMPT_NONE"] == projected["value"]
    assert constants["SIM_ATTEMPT_NONE"] == projected["value"]
    # EVERY OTHER WRITE IS A CLEAR, A BLANKED MARKER OR A RESTORE FROM THE
    # CARRIER. Nothing else reaches a cell.
    for owner, clear, restore in OWNERS:
        for procedure in (clear, restore):
            for assignment in re.findall(r"\.Value2 = (.+)", _procedure(owner, procedure)):
                assert assignment.strip() in (
                    "vbNullString", "CALC_ATTEMPT_NONE", "SIM_ATTEMPT_NONE",
                    "carried(1)"), assignment


def test_18_no_counter_is_advanced_and_no_identity_is_allocated() -> None:
    """§2, §8 and §14. A reset that renumbered anything would let a discarded run
    be re-created by a future one."""
    bodies = [_reset().code] + [_addition(owner) for owner, _c, _r in OWNERS]
    for body in bodies:
        for banned in ("SimNonceAllocate", "NextId", "AllocateAutoNonce",
                       "CandidateRunId", "FinalCommit"):
            assert banned not in body, banned


# ===========================================================================
# E. THE ORCHESTRATOR HOLDS NO GEOMETRY
# ===========================================================================
def test_19_mod_reset_names_no_address_and_no_layout_constant() -> None:
    """§4: "Do NOT make modReset a second authority for bank addresses, stamp
    addresses, fingerprint layouts or result record geometry." """
    code = _reset().code
    assert not re.search(r'"\s*\$?[A-Z]{1,3}\$?\d+', code), "modReset spells an address"
    referenced = _reset().referenced_upper_identifiers
    generated = set(_generated_constants())
    assert not (referenced & generated), sorted(referenced & generated)
    for banned in ("SIM_BANK", "TBL_CALC", "_ROW_", "_COLUMN", "_RANGE"):
        assert banned not in code, banned


def test_20_mod_reset_calls_exactly_the_eight_owner_procedures() -> None:
    """§4. Four owners, two calls each, and nothing else reaches a store."""
    code = _reset().code
    calls = set(re.findall(r"\bmod\w+\.(\w+)", code))
    expected = {name for _o, clear, restore in OWNERS for name in (clear, restore)}
    owned = {name for name in calls if name.endswith("Publication")}
    assert owned == expected, sorted(owned ^ expected)
    for owner, clear, restore in OWNERS:
        assert f"{owner}.{clear}" in code, clear
        assert f"{owner}.{restore}" in code, restore


def test_21_the_undo_carrier_is_opaque_to_the_orchestrator() -> None:
    """§4. It is built by an owner and read by that owner. modReset never looks
    inside one, so it cannot come to depend on what a store's shape is."""
    code = _reset().code
    for carrier in ("calcUndo", "simUndo", "annualUndo", "sensitivityUndo"):
        assert carrier in code, carrier
        assert f"{carrier}(" not in code, f"modReset indexes {carrier}"
    assert "IsEmpty" not in code, (
        "modReset decides which owners to restore; every owner is asked")


def test_22_the_restore_order_is_the_reverse_of_the_clear_order() -> None:
    """§5. All earlier clears come back, and they come back in the order that
    undoes them."""
    clears = _procedure("modReset", "ClearEveryPublication")
    rollback = _procedure("modReset", "RollbackAndReport")
    forward = [owner for owner, _c, _r in OWNERS if f"{owner}." in clears]
    backward = [owner for owner, _c, _r in OWNERS if f"{owner}." in rollback]
    assert forward == [owner for owner, _c, _r in OWNERS]
    assert backward == forward
    positions = [rollback.index(f"{owner}.") for owner in backward]
    assert positions == sorted(positions, reverse=True), (
        "the rollback does not run in reverse order")


# ===========================================================================
# F. CONFIRMATION, REPORTING AND THE ENVELOPE
# ===========================================================================
def test_23_the_destructive_confirmation_is_the_accepted_one() -> None:
    """§6. One confirmation mechanism, and the harness can drive it."""
    code = _reset().code
    assert "modAppState.AskConfirm(ConfirmationSummary(), True)" in code
    assert code.count("AskConfirm") == 1
    assert "MsgBox" not in code, "a second dialog is a second reporting mechanism"
    assert "ConfirmDestructiveChange" not in code, (
        "the confirmation goes through AskConfirm so automation can answer it")
    assert "gAutomationActive" not in code, (
        "the automation gate belongs to the reporting owner, not to a command")


def test_24_a_cancellation_changes_nothing_and_is_not_a_failure() -> None:
    """§6 and §13-E. The confirmation is asked BEFORE any owner is called, so a
    cancellation needs no rollback; an empty success message is how the accepted
    reporting owner says nothing happened."""
    body = _procedure("modReset", "ResetResults")
    confirm = body.index("AskConfirm")
    assert confirm < body.index("ClearEveryPublication")
    cancelled = body[confirm:body.index("ClearEveryPublication")]
    assert "modAppState.Succeeded(vbNullString)" in cancelled
    assert "Failed" not in cancelled, "a cancellation is not an error"


def test_25_the_outcome_goes_through_the_accepted_reporting_owner() -> None:
    """§7. Announce records for automation always and shows a dialog only when a
    human is there. Reset adds no reporting of its own."""
    endpoint = _procedure("modReset", ENDPOINT)
    assert endpoint.count("modAppState.Announce") == 4, endpoint.count("Announce")
    assert "modAppState.CaptureAppState()" in endpoint
    assert "modAppState.FinishOperation(state)" in endpoint
    for label in re.findall(r'modAppState\.Failed\("([^"]+)"', _reset().raw):
        assert label == "Reset Results", label


def test_26_the_success_sentence_exists_in_exactly_one_place() -> None:
    """§7: "Do not duplicate this sentence in multiple modules/specs." """
    sentence = "Results reset. Model inputs and identity counters were preserved."
    assert _reset().raw.count(sentence) == 1
    for path in list(SRC.glob("*.bas")) + [SRC / "ThisWorkbook.vba"]:
        if path.name != "modReset.bas":
            assert sentence not in path.read_text(encoding="utf-8"), path.name
    for path in sorted(SPEC.glob("*.yaml")):
        assert sentence not in path.read_text(encoding="utf-8"), path.name


def test_27_a_failed_reset_reports_the_restore_outcome() -> None:
    """§7. The user has to be told whether what was cleared came back."""
    rollback = _procedure("modReset", "RollbackAndReport")
    assert "was put back" in _reset().raw
    assert "could not be fully restored" in _reset().raw
    assert "If Len(problems) = 0 Then" in rollback


def test_28_a_second_reset_is_a_fixpoint() -> None:
    """§8. Every write is a blank or a contract initial, and nothing branches on
    what is already there - so running it twice reaches the same place and the
    second run cannot fail merely because the first succeeded."""
    for owner, clear, _restore in OWNERS:
        body = _procedure(owner, clear)
        assert "If " not in body, f"{clear} branches on the state it is clearing"


def test_29_reset_never_unprotects_a_sheet() -> None:
    """§10. Protection is UserInterfaceOnly, so every write below already works.
    An Unprotect here would leave the workbook open if the command then failed.
    """
    bodies = [_reset().code] + [_addition(owner) for owner, _c, _r in OWNERS]
    for body in bodies:
        for banned in ("Unprotect", ".Protect", "modProtection"):
            assert banned not in body, banned


# ===========================================================================
# H. THE BUTTON
# ===========================================================================
def test_30_the_command_is_declared_bound_and_reachable() -> None:
    """§9. The fifth of the six commands the 6ab8f6a contract authorises."""
    structure = _structure()
    buttons = {b.shape_name: b for b in structure.buttons}
    assert SHAPE in buttons
    button = buttons[SHAPE]
    assert (button.caption, button.entry_point, button.sheet) == (
        CAPTION, ENDPOINT, "Setup")
    assert ENDPOINT in structure.entry_points
    assert f"Public Sub {ENDPOINT}()" in _reset().raw
    # AND IT IS NOT AN INTEGRATOR API. A destructive command that could be driven
    # without a confirmation is not one an integrator should be handed.
    assert ENDPOINT not in structure.api_procedures


def test_31_the_button_lands_on_the_next_anchor_and_moves_none_of_the_nine() -> None:
    """§9. "Do not alter the existing nine buttons already accepted."

    RESTATED AT P10-UX. The block now LEADS with Apply / Update Timeline, which
    the visual review moved out of the Applied Timeline block it was drawn over,
    so every command sits one pitch lower than it did. Nothing about Reset
    Results changed: it is still the last command in the block, still on the
    declared column at the declared pitch, still bound to the same endpoint. The
    claim is now made by POSITION IN THE BLOCK rather than by a row number,
    which is what it always meant.
    """
    structure = _structure()
    block = structure.commands["block"]
    column = str(block["button_column"])
    first = int(block["first_button_row"])
    pitch = int(block["button_row_pitch"])
    commands = [b for b in structure.buttons if b.anchor_cell.startswith(column)
                and int(b.anchor_cell[len(column):]) >= first]
    anchors = [b.anchor_cell for b in commands]
    assert anchors == [f"{column}{first + pitch * i}" for i in range(len(commands))]
    assert len(structure.buttons) == 10, len(structure.buttons)
    # LAST IN THE BLOCK, wherever the block starts. A destructive command that
    # crept up among the run commands would be a different mistake, and the
    # index below is what refuses it.
    assert commands[-1].shape_name == SHAPE
    assert buttons_at(structure, f"{column}{first + pitch * (len(commands) - 1)}") == SHAPE


def buttons_at(structure, anchor: str) -> str:
    for button in structure.buttons:
        if button.anchor_cell == anchor:
            return button.shape_name
    raise AssertionError(f"no button at {anchor}")


def test_32_repair_profiling_has_not_started() -> None:
    """§9 and the batch fence: the sixth command belongs to a later step."""
    assert not (SRC / "modRepair.bas").exists()
    structure = (SPEC / "structure_contract.yaml").read_text(encoding="utf-8")
    assert "PCCM_RepairProfiling" not in structure
    assert "Repair" not in _reset().raw


# ===========================================================================
# I. THE STATIC SCENARIOS
#
# Each names the §13 scenario it stands for and asserts the specific thing that
# scenario is about, from the source and the contracts. None of them runs Excel
# and none of them claims to.
# ===========================================================================
def test_33_scenario_a_a_calculated_model_loses_its_calculation_only() -> None:
    """A. The calculation publication goes; the live reading is derived."""
    assert "CALC_STATE_VALUE_RANGE" in _procedure("modCalcReport", "PublicationBlocks")
    # THE LIVE READING IS NOT WRITTEN BY THE RESET. Model Check asks the owner's
    # pure evaluator, which reads NOT CALCULATED from a blank fingerprint on a
    # valid model and INVALID on an invalid one.
    derive = _procedure("modCalcReport", "DeriveStatus")
    assert "CALC_STATUS_NOT_CALCULATED" in derive and "CALC_STATUS_INVALID" in derive
    assert "CalcReportDerivedStatus" in _module("modResultsState").code


def test_34_scenario_b_all_four_publication_layers_are_cleared() -> None:
    """B. Calculation, simulation, annual and sensitivity - in that order, each
    through its own owner."""
    body = _procedure("modReset", "ClearEveryPublication")
    order = [owner for owner, _c, _r in OWNERS]
    positions = [body.index(f"{owner}.") for owner in order]
    assert positions == sorted(positions), order


def test_35_scenario_c_an_invalid_model_stays_invalid() -> None:
    """C. Reset touches no input, so nothing that made the model invalid moves;
    the owner derives INVALID from the inputs exactly as it did before."""
    assert "PrepareCurrentCalculation" not in _addition("modCalcReport")
    assert "CalcPrepareSimulationInputs" not in _reset().code


def test_36_scenario_f_and_g_every_earlier_clear_is_restored() -> None:
    """F and G. A failure between any two owners restores every owner that ran,
    and the failpoints sit exactly at those boundaries."""
    body = _procedure("modReset", "ClearEveryPublication")
    stages = re.findall(r"modAppState\.FailPointCheck (FAILPOINT_RESET_\w+)", body)
    assert stages == ["FAILPOINT_RESET_CALCULATION", "FAILPOINT_RESET_SIMULATION",
                      "FAILPOINT_RESET_ANNUAL", "FAILPOINT_RESET_SENSITIVITY"], stages
    for stage in stages:
        assert f"Public Const {stage} As String" in _reset().raw, stage
    # AND NO OWNER CARRIES ONE OF ITS OWN, so a failpoint can only fire between
    # two owners - which is the boundary the scenarios name.
    for owner, _c, _r in OWNERS:
        assert "FailPointCheck" not in _addition(owner), owner


def test_37_scenario_h_the_ownership_design_is_protection_compatible() -> None:
    """H. Every write is a code write to a UserInterfaceOnly-protected sheet, and
    no shape, name, table structure or sheet visibility is touched."""
    bodies = [_addition(owner) for owner, _c, _r in OWNERS]
    for body in bodies:
        for banned in ("Shapes", "Visible", "ListObjects.Add", "Names.Add", "Delete"):
            assert banned not in body, banned


def test_38_scenario_i_identity_continuity_is_untouched_by_construction() -> None:
    """I. Proved by test_11 for the simulation rows and test_13 for the permanent
    ids; this asserts the one remaining route - that no owner spells a worksheet
    address of its own anywhere in the block."""
    for owner, _clear, _restore in OWNERS:
        addition = _addition(owner)
        assert re.search(r"\.ClearContents", addition), owner
        assert not re.search(r'"\$?[A-Z]{1,3}\$?\d+', addition), owner


def test_39_scenario_j_no_presentation_sheet_needs_a_phase_10_special_case() -> None:
    """J. Model Check, Results and the Dashboard read the same owners they always
    did. Nothing in this batch touches a presentation sheet or an adapter."""
    presentation = {name for name, value in _generated_constants().items()
                    if value in ("Dashboard", "Model Check", "Results",
                                 "Sensitivity", "Methodology")}
    assert presentation, "no presentation sheet constant was found to check"
    for body in [_reset().code] + [_addition(o) for o, _c, _r in OWNERS]:
        for constant in presentation:
            assert constant not in body, constant
        assert "modResultsState" not in body


def test_40_the_projection_is_addresses_only() -> None:
    """§17. Where a later run must look. No expected value, no reset, no claim
    that one has ever been executed."""
    projection = _projection()
    assert projection["command"] == ENDPOINT
    assert set(projection) == {"command", "preserved", "publications"}
    def leaves(node, path=""):
        if isinstance(node, dict):
            for key, value in node.items():
                assert "expected" not in key.lower(), f"{path}/{key}"
                yield from leaves(value, f"{path}/{key}")
        elif isinstance(node, list):
            for index, value in enumerate(node):
                yield from leaves(value, f"{path}[{index}]")
        else:
            yield path, node

    for path, value in leaves(projection):
        # ADDRESSES, NAMES AND ROW NUMBERS. A float would be a quantity, and a
        # quantity in a projection is an expected value by another name.
        assert isinstance(value, (str, int)) and not isinstance(value, bool), (
            path, value)


# THE EXACT CLEAR SCOPE, DECLARED ONCE.
#
# Each owner answers "what is my publication" with ONE list, and the list below
# is that answer written out. A block the owner adds without it being reviewed
# here fails; a block it drops fails too. Everything else about the transaction
# follows from the shape rather than from a second list: the clear walks the
# whole list, the capture is one entry per block, and the restore walks the
# whole carrier - so "everything cleared is restored" is structural rather than
# something a table has to keep in step.
PUBLICATION_BLOCKS = {
    "modCalcReport": [
        "BodyAddress(TBL_CALC_YEARS)", "BodyAddress(TBL_CALC_INFLATION_FACTORS)",
        "BodyAddress(TBL_CALC_FX)", "BodyAddress(TBL_CALC_DRIVERS)",
        "BodyAddress(TBL_CALC_ANNUAL)",
        "CALC_TOTALS_VALUE_RANGE", "CALC_STATE_VALUE_RANGE",
    ],
    "modSimReport": [
        "SnapshotRange(SIM_BANK_A)", "SnapshotRange(SIM_BANK_B)",
        "SummaryRange(SIM_BANK_A)", "SummaryRange(SIM_BANK_B)",
        "ContingencyRange(SIM_BANK_A)", "ContingencyRange(SIM_BANK_B)",
        "IterationRange(SIM_BANK_A, 0, SIM_MAX_ITERATIONS)",
        "IterationRange(SIM_BANK_B, 0, SIM_MAX_ITERATIONS)",
        "PublicationRecordRange()",
    ],
    "modSimAnnualStore": [
        "StampAddress(SIM_BANK_A)", "StampAddress(SIM_BANK_B)",
        "BlockAddress(SIM_BANK_A)", "BlockAddress(SIM_BANK_B)",
    ],
    "modSimPostReport": [
        "StampAddress(SIM_BANK_A)", "StampAddress(SIM_BANK_B)",
        "RecordAddress(SIM_BANK_A)", "RecordAddress(SIM_BANK_B)",
    ],
}


@pytest.mark.parametrize("owner,clear,restore", OWNERS)
def test_41_the_publication_is_one_declared_list(owner: str, clear: str,
                                                 restore: str) -> None:
    """§3 and §18-3. The exact clear scope, and it is exactly one list."""
    blocks = _procedure(owner, "PublicationBlocks")
    entries = re.findall(r"^\s*(?:PublicationBlocks = Array\( _|)?(.*)$", blocks, re.M)
    for expected in PUBLICATION_BLOCKS[owner]:
        assert expected in blocks, f"{owner} does not publish through {expected}"
    # NOTHING ELSE IS IN IT. Counting the separators is what makes this exact
    # rather than "contains at least".
    commas = blocks.count(",") - blocks.count("SIM_BANK_A, 0,") * 2 \
        - blocks.count("SIM_BANK_B, 0,") * 2
    assert commas == len(PUBLICATION_BLOCKS[owner]) - 1, (owner, commas)


@pytest.mark.parametrize("owner,clear,restore", OWNERS)
def test_42_the_clear_and_the_restore_walk_the_whole_list(owner: str, clear: str,
                                                          restore: str) -> None:
    """§5 and §14. THIS IS WHERE "rollback restores the outputs but not the
    attempt history" DIES.

    The clear captures one entry per declared block, clears every declared block,
    and the restore walks every entry of the carrier. Nothing selects a subset at
    any of the three points, so a publication cannot be cleared without being
    captured and cannot be captured without being put back.
    """
    clearing = _procedure(owner, clear)
    restoring = _procedure(owner, restore)
    assert "addresses = PublicationBlocks()" in clearing
    assert "ReDim captured(LBound(addresses) To UBound(addresses))" in clearing
    assert clearing.count("For index = LBound(addresses) To UBound(addresses)") == 2, (
        f"{clear} does not both capture and clear over the whole list")
    assert "captured(index) = CapturedBlock(CStr(addresses(index)))" in clearing
    assert ".Range(CStr(addresses(index))).ClearContents" in clearing
    # THE WHOLE CARRIER, TO THE END OF THE LINE. A bound that is one short - or
    # one that starts one late - restores every publication but the last, which
    # for the calculation store is C13:C20: the attempt history.
    assert re.search(r"^\s*For index = (?:LBound\(undo\) To UBound\(undo\)|"
                     r"UBound\(undo\) To LBound\(undo\) Step -1)\s*$",
                     restoring, re.M), restore
    assert "RestoredBlock undo(index)" in restoring
    # AND THE CAPTURE IS BOUNDED BY THE SHEET, NOT BY THE CEILING. A million-row
    # snapshot to put blanks back would be a backup of the sheet.
    captured = _procedure(owner, "CapturedBlock")
    assert "Application.Intersect(" in captured and "UsedRange" in captured
    assert "If block Is Nothing Then Exit Function" in captured


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
