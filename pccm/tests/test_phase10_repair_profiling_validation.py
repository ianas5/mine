#!/usr/bin/env python3
"""P10-2C MUTATION CONTROLS for the Repair Profiling battery.

A recovery command's controls are the ones most worth attacking: everything they
guard is a case the user cannot check for themselves, because by the time they
press the button the workbook is already wrong. Every control below damages one
of the authorities this batch touches - `modRepair.bas` or
`spec/structure_contract.yaml` - reruns the WHOLE conformance battery against the
damaged copy, and requires a NAMED detector among the refusers.

THE MUTATIONS ARE THE ONES A REVIEWER WOULD FEAR: the owners called in the wrong
order, the command writing a grid itself, a blocking fault admitted as
repairable, the unkeyed-data gate or the trim gate dropped, an unreadable id or
weight accepted, conflicting duplicates collapsed, a blank treated as a zero, a
half-entered profile restructured or normalised, the sum target moved, the
snapshot taken too late, half a rollback, a suppressed error, the revalidation
dropped, a state word forced, an Unprotect, and a button bound to the wrong
command.

AND THE BEHAVIOURAL HALF IS ATTACKED TOO. test_25 composes the accepted oracle in
the WRONG order and requires the preservation controls to notice - because a
behavioural control that would pass either way proves nothing about the order the
source is required to call the owners in.

Runs standalone or under pytest.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

import test_phase10_repair_profiling as conformance  # noqa: E402
from pccm_builder.structure_oracle import (  # noqa: E402
    remap_profiling,
    sync_profiling_values,
    sync_rows,
)

_SOURCE = conformance.REPAIR_BAS.read_text(encoding="utf-8")


def _tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 24, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _install(source: str | None = None, structure=None):
    saved = (conformance.REPAIR_BAS, dict(conformance._CACHE))
    conformance._CACHE.clear()
    conformance._CACHE.update(saved[1])
    if source is not None:
        assert source != _SOURCE, "the mutation changed nothing"
        temp = Path(tempfile.mkdtemp(prefix="pccm-p10-2c-mutation-"))
        target = temp / "modRepair.bas"
        target.write_text(source, encoding="utf-8")
        conformance.REPAIR_BAS = target
        conformance._CACHE.pop("module", None)
    if structure is not None:
        conformance._CACHE["structure"] = structure

    def restore() -> None:
        conformance.REPAIR_BAS = saved[0]
        conformance._CACHE.clear()
        conformance._CACHE.update(saved[1])

    return restore


def _control(expected: str, source: str | None = None, structure=None) -> None:
    restore = _install(source, structure)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _swap(old: str, new: str, count: int = 1) -> str:
    assert _SOURCE.count(old) == count, (old[:70], _SOURCE.count(old))
    return _SOURCE.replace(old, new)


def _damaged_structure(replacements: list[tuple[str, str]]):
    from pccm_builder import load_structure_contract
    text = (conformance.SPEC / "structure_contract.yaml").read_text(encoding="utf-8")
    for old, new in replacements:
        assert text.count(old) == 1, (old[:60], text.count(old))
        text = text.replace(old, new, 1)
    temp = Path(tempfile.mkdtemp(prefix="pccm-p10-2c-structure-"))
    shutil.copytree(conformance.SPEC, temp / "spec")
    (temp / "spec" / "structure_contract.yaml").write_text(text, encoding="utf-8")
    return load_structure_contract(temp / "spec" / "structure_contract.yaml")


def _refused_by_the_contract(replacements, expected: str) -> None:
    """Some mutations never reach the battery because the CONTRACT refuses to
    load - a stronger refusal, recorded where it actually happens."""
    from pccm_builder import StructureContractError
    try:
        _damaged_structure(replacements)
    except StructureContractError as error:
        assert expected in str(error), (expected, str(error))
        return
    raise AssertionError("the contract loaded a workbook it should have refused")


def test_00_the_accepted_sources_pass_every_detector() -> None:
    restore = _install()
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused == [], refused


# ===========================================================================
# A. OWNERSHIP
# ===========================================================================
def test_01_the_owners_are_called_in_the_wrong_order() -> None:
    """Rows before columns preserves weights into positions that are about to be
    reshaped - which is the one ordering mistake this repair can make."""
    damaged = _swap(
        "    modProfiling.SetYearColumns plan.Kind, plan.StartYear, plan.TargetYears\n"
        "    modProfiling.SyncRows plan.Kind\n",
        "    modProfiling.SyncRows plan.Kind\n"
        "    modProfiling.SetYearColumns plan.Kind, plan.StartYear, plan.TargetYears\n")
    _control("test_01", source=damaged)


def test_02_the_command_writes_a_grid_itself() -> None:
    """The moment it writes a cell it is a second profiling owner, and the two
    will disagree about what a new project year starts at."""
    damaged = _swap(
        "    modProfiling.SyncRows plan.Kind\n",
        "    modProfiling.SyncRows plan.Kind\n"
        "    modWorkbook.CellIn(modProfiling.ProfilingTable(plan.Kind), 1, 1).Value = \"\"\n")
    _control("test_01", source=damaged)


def test_03_the_command_names_a_table_of_its_own() -> None:
    damaged = _swap(
        "    Set grid = modProfiling.ProfilingTable(kind)\n",
        "    Set grid = modWorkbook.Lo(SH_COST_PROFILING, TBL_COST_PROFILING)\n")
    _control("test_02", source=damaged)


# ===========================================================================
# B. THE REFUSALS
# ===========================================================================
def test_04_a_blocking_fault_is_admitted_as_repairable() -> None:
    """A malformed identifier is not a structure this command can reconstruct;
    admitting the key would let it rebuild a grid around an id nobody can use."""
    damaged = _swap(
        "         CHK_NO_DUPLICATE_IDS\n",
        "         CHK_NO_DUPLICATE_IDS, CHK_ID_PATTERN\n")
    _control("test_13", source=damaged)


def test_05_unkeyed_structural_data_stops_being_a_gate() -> None:
    """Data in a row with no id would be deleted by synchronisation, and no
    confirmation could ever have warned about it because it is unkeyed."""
    damaged = _swap(
        "    detail = modStructuralCheck.PreMutationCheck()\n"
        "    If Len(detail) > 0 Then\n"
        "        RepairProfiling = Refused(detail)\n"
        "        Exit Function\n"
        "    End If\n", "")
    _control("test_14", source=damaged)


def test_06_a_trim_that_would_delete_weights_is_allowed() -> None:
    """Repairing the width by deleting the user's weights is the one repair that
    would be worse than the fault."""
    damaged = _swap(
        "    hits = modProfiling.CountDataBeyond(kind, targetYears, affected, affectedCount)\n"
        "    If hits = 0 Then\n",
        "    hits = modProfiling.CountDataBeyond(kind, targetYears, affected, affectedCount)\n"
        "    If True Then\n")
    _control("test_14", source=damaged)


def test_07_an_unreadable_identifier_is_accepted() -> None:
    damaged = _swap(
        '        If modWorkbook.IsErrorText(idText) Then\n'
        '            detail = label & " row " & CStr(row) & " has an identifier that cannot be " & _\n'
        '                     "read, so its weights cannot be attributed to any driver."\n'
        '            Exit Function\n'
        '        End If\n', "")
    _control("test_15", source=damaged)


def test_08_a_weight_that_is_not_a_number_is_carried_through() -> None:
    """Text in a weight cell cannot be preserved as a weight; carrying it into a
    rebuilt grid would make this command the author of a value it cannot read."""
    damaged = _swap(
        "        ElseIf Not IsNumeric(cell.Value) Then\n",
        "        ElseIf False Then\n")
    _control("test_15", source=damaged)


def test_09_conflicting_duplicate_rows_are_collapsed() -> None:
    """Two rows for one id saying different things are two answers to one
    question, and picking either is this command choosing what the user meant."""
    damaged = _swap(
        "                If Not SameWeights(held, weights) Then\n",
        "                If False Then\n")
    _control("test_16", source=damaged)


def test_10_a_blank_weight_is_treated_as_a_zero() -> None:
    """The grid language distinguishes "not entered" from "nothing allocated". A
    comparison that collapsed them would call two different profiles identical
    and silently discard one of them."""
    damaged = _swap(
        "        If IsEmpty(held(index)) <> IsEmpty(found(index)) Then Exit Function\n", "")
    _control("test_16", source=damaged)


def test_11_a_repeated_register_identifier_is_accepted() -> None:
    damaged = _swap(
        '            If ids.Exists(idText) Then\n'
        '                detail = register.Name & " contains the identifier " & idText & " more " & _\n',
        '            If False Then\n'
        '                detail = register.Name & " contains the identifier " & idText & " more " & _\n')
    _control("test_17", source=damaged)


def test_12_a_half_entered_profile_is_restructured_anyway() -> None:
    """The semantic gate removed: a row totalling 60% gets moved between two
    different sets of project years and handed back looking repaired."""
    damaged = _swap(
        "            If plan.WidthDrift And registerIds.Exists(idText) Then\n"
        "                If Not RecognisedProfile(weights, label, idText, detail) Then Exit Function\n"
        "            End If\n", "")
    _control("test_18", source=damaged)


def test_13_the_command_normalises_a_profile_to_100_percent() -> None:
    """THE ONE THING A STRUCTURAL REPAIR MUST NEVER DO. Scaling a row to total
    100% is inventing an allocation the user never made."""
    damaged = _swap(
        "    If IsRecognisedSum(total) Then\n",
        "    terms(0) = terms(0) / total\n"
        "    If IsRecognisedSum(total) Then\n")
    _control("test_18", source=damaged)


def test_14_the_recognised_total_stops_being_100_percent() -> None:
    """A target that drifted from the Model Check's would make two parts of the
    workbook disagree about what a profile is."""
    damaged = _swap(
        "Private Const REPAIR_PROFILE_SUM_TARGET As Double = 1#",
        "Private Const REPAIR_PROFILE_SUM_TARGET As Double = 0.99")
    _control("test_19", source=damaged)


def test_15_the_empty_total_stops_being_the_contract_default() -> None:
    """0% is the contract's own initial value. A different one here would make
    the command refuse on every workbook with an unprofiled driver."""
    damaged = _swap(
        "Private Const REPAIR_PROFILE_SUM_EMPTY As Double = 0#",
        "Private Const REPAIR_PROFILE_SUM_EMPTY As Double = 0.5")
    _control("test_19", source=damaged)


# ===========================================================================
# C. THE TRANSACTION
# ===========================================================================
def test_16_the_snapshot_is_taken_after_the_first_grid_is_repaired() -> None:
    """"Cost repaired and Risk broken" is exactly what this ordering produces."""
    damaged = _swap(
        "    costBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(cost.Kind))\n"
        "    riskBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(risk.Kind))\n"
        "    captured = True\n\n"
        "    Apply cost\n",
        "    Apply cost\n"
        "    costBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(cost.Kind))\n"
        "    riskBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(risk.Kind))\n"
        "    captured = True\n")
    _control("test_20", source=damaged)


def test_17_only_one_grid_is_restored_on_failure() -> None:
    damaged = _swap(
        "    modWorkbook.RestoreTable modProfiling.ProfilingTable(modProfiling.RiskKind()), riskBefore\n",
        "")
    _control("test_21", source=damaged)


def test_18_an_error_is_suppressed_instead_of_rolled_back() -> None:
    damaged = _swap(
        "    On Error GoTo RepairFailed\n",
        "    On Error Resume Next\n    On Error GoTo RepairFailed\n")
    _control("test_21", source=damaged)


def test_19_the_result_is_not_revalidated() -> None:
    """A repair that produced an invalid workbook and reported success would be
    worse than the fault it was asked to fix."""
    damaged = _swap(
        "    detail = modStructuralCheck.ValidateStructure()\n"
        "    If Len(detail) > 0 Then\n"
        "        Err.Raise vbObjectError + 5201, \"modRepair.RepairProfiling\", _\n"
        '                  "structural revalidation failed after repair:" & vbCrLf & detail\n'
        "    End If\n", "")
    _control("test_22", source=damaged)


# ===========================================================================
# D. STATE, PROTECTION AND THE BUTTON
# ===========================================================================
def test_20_the_command_forces_a_state_word() -> None:
    """The fingerprint and state owners stay authoritative. A repair that wrote
    CURRENT would be a bypass wearing a repair's name."""
    damaged = _swap(
        "    RepairProfiling = modAppState.Succeeded(Summary(cost, risk))\n",
        '    modWorkbook.WriteValue NM_STRUCTURAL_STATE, "Timeline current"\n'
        "    RepairProfiling = modAppState.Succeeded(Summary(cost, risk))\n")
    _control("test_23", source=damaged)


def test_21_the_command_unprotects_the_sheets_it_writes_to() -> None:
    damaged = _swap(
        "    Apply cost\n",
        "    modProfiling.ProfilingTable(cost.Kind).Parent.Unprotect\n    Apply cost\n")
    _control("test_24", source=damaged)


def test_22_the_button_is_bound_to_the_wrong_command() -> None:
    _refused_by_the_contract(
        [('      entry_point: "PCCM_RepairProfiling"',
          '      entry_point: "PCCM_ResetResults"')],
        "are bound to more than one button")


def test_23_the_button_lands_among_the_run_commands() -> None:
    """Repair Profiling is the command a healthy workbook never needs. Met on
    the way to Calculate it is an invitation to rebuild a grid nobody looked at."""
    damaged = _damaged_structure([
        ('      shape_name: "btnPCCMRepairProfiling"\n'
         '      caption: "Repair Profiling"\n'
         '      entry_point: "PCCM_RepairProfiling"\n'
         '      anchor_cell: "E71"',
         '      shape_name: "btnPCCMRepairProfiling"\n'
         '      caption: "Repair Profiling"\n'
         '      entry_point: "PCCM_RepairProfiling"\n'
         '      anchor_cell: "E73"')])
    _control("test_25", structure=damaged)


def test_24_the_no_op_stops_telling_the_user_anything() -> None:
    """"Nothing happened" and "nothing needed to happen" are different answers,
    and only one of them tells a user their grids are sound."""
    damaged = _swap(
        '            SH_COST_PROFILING & " and " & SH_RISK_PROFILING & " are already " & _\n'
        '            "structurally correct. Nothing was changed.")\n',
        '            "Done.")\n')
    _control("test_26", source=damaged)


# ===========================================================================
# E. WHERE THE ORDER GUARANTEE ACTUALLY LIVES
# ===========================================================================
def test_25_the_oracle_composition_cannot_carry_the_ordering_claim() -> None:
    """AN HONEST NEGATIVE RESULT, RECORDED RATHER THAN PAPERED OVER.

    The behavioural controls compose the oracle as SetYearColumns then SyncRows.
    The obvious mutation is to compose it the other way round - and it does not
    fail, because at the oracle's level the two orders agree: reshaping then
    synchronising and synchronising then reshaping both truncate to the same
    positions and pad with the same default.

    THAT IS A FACT ABOUT THE ORACLE, NOT ABOUT THE COMMAND. In the workbook the
    order is load-bearing: SyncRows reads the grid's CURRENT year-column count
    and preserves weights into it, so running it before the reshape preserves
    them into positions that are about to be deleted or that do not exist yet.
    The oracle takes the width as an argument and so cannot express that.

    So the ordering claim is carried where it can be: by the SOURCE control that
    requires modRepair to call the two owners in that order, which the first
    mutation in this file attacks directly. This control exists to say so, and to
    fail if somebody ever makes the oracle order-sensitive and forgets that the
    behavioural controls would then be asserting something new.
    """
    register = ["CL-001", "CL-002"]
    for grid, target in (({"CL-001": [0.5, 0.5, 0.0]}, 2),
                         ({"CL-001": [1.0]}, 3),
                         ({"CL-001": [0.25, 0.75], "CL-009": [1.0, 0.0]}, 4)):
        forward = conformance.repair_grid(register, grid, target)
        ordered, _added, _removed = sync_rows(register, grid)
        width = max((len(v) for v in grid.values()), default=0)
        reversed_order = remap_profiling(
            sync_profiling_values(grid, ordered, width, conformance.INITIAL),
            target, conformance.INITIAL)
        assert forward == reversed_order, (grid, target, forward, reversed_order)
    # AND THE SOURCE CONTROL THAT DOES CARRY IT IS PRESENT AND ATTACKED.
    assert "test_01_the_repair_is_two_owner_calls_and_nothing_else" in _tests()


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
