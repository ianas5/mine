#!/usr/bin/env python3
"""P10-2B MUTATION CONTROLS for the Reset Results battery.

A conformance test that cannot fail proves nothing. Every control below damages
one of the authorities this step touches - `modReset.bas`, one of the four
publication owners, `spec/structure_contract.yaml` or the emitted projection -
reruns the WHOLE conformance battery against the damaged copy, and requires a
NAMED detector among the refusers.

THE SIXTEEN ARE THE SIXTEEN THE AUTHORISATION LISTS, in its order. Each is the
regression a reviewer would actually fear from a destructive command: a reset
that reaches an input, that renumbers an identity, that clears one bank and not
the other, that leaves a layer published, that forces a state word, that mutates
after a cancellation, that leaves a half-reset workbook, or that quietly
unprotects the sheets it writes to.

Nothing here writes to the repository: damaged copies live in a temporary
directory and the conformance module is pointed at them for one control.

Runs standalone or under pytest.
"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

import test_phase10_reset_results as conformance  # noqa: E402

_SOURCES = {name: (conformance.SRC / f"{name}.bas").read_bytes().decode("utf-8")
            for name in ("modReset", "modCalcReport", "modSimReport",
                         "modSimAnnualStore", "modSimPostReport")}
_STRUCTURE = (conformance.SPEC / "structure_contract.yaml").read_text(encoding="utf-8")
_PROJECTION = conformance.PROJECTION.read_text(encoding="utf-8")


def _tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 35, names
    return names


def _run_battery() -> list[str]:
    """Every control, including the parametrised ones, run over all four owners."""
    refused: list[str] = []
    for name in _tests():
        function = getattr(conformance, name)
        marks = getattr(function, "pytestmark", [])
        cases: list[tuple] = [()]
        for mark in marks:
            if mark.name == "parametrize":
                cases = [tuple(case) for case in mark.args[1]]
        for case in cases:
            try:
                function(*case)
            except BaseException:  # noqa: BLE001 - any refusal counts
                refused.append(name)
                break
    return refused


def _install(sources: dict[str, str] | None = None, structure: str | None = None,
             projection: str | None = None):
    saved = (conformance.SRC, conformance.SPEC, conformance.PROJECTION,
             dict(conformance._CACHE))
    conformance._CACHE.clear()
    temp = Path(tempfile.mkdtemp(prefix="pccm-p10-2b-mutation-"))
    if sources:
        src = temp / "vba"
        shutil.copytree(saved[0], src)
        for name, text in sources.items():
            assert text != _SOURCES[name], f"the mutation changed nothing in {name}"
            (src / f"{name}.bas").write_bytes(text.encode("utf-8"))
        conformance.SRC = src
    if structure is not None:
        assert structure != _STRUCTURE, "the mutation changed nothing"
        spec = temp / "spec"
        shutil.copytree(saved[1], spec)
        (spec / "structure_contract.yaml").write_text(structure, encoding="utf-8")
        conformance.SPEC = spec
    if projection is not None:
        assert projection != _PROJECTION, "the mutation changed nothing"
        target = temp / "phase10_reset_inspection.json"
        target.write_text(projection, encoding="utf-8")
        conformance.PROJECTION = target

    def restore() -> None:
        conformance.SRC, conformance.SPEC, conformance.PROJECTION = saved[:3]
        conformance._CACHE.clear()
        conformance._CACHE.update(saved[3])

    return restore


def _control(expected: str, sources: dict[str, str] | None = None,
             structure: str | None = None, projection: str | None = None) -> None:
    restore = _install(sources, structure, projection)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _swap(name: str, old: str, new: str, count: int = 1) -> dict[str, str]:
    """LINE-ENDING AWARE, AND THAT IS NOT A DETAIL.

    modCalcReport.bas is CRLF and every other module is LF. An anchor written
    with LF and matched against CRLF finds nothing, `count` would be zero, and a
    control whose mutation was never planted PASSES for the wrong reason - the
    single most dangerous shape a mutation control can take. So the anchor is
    translated into the file's own ending before it is matched.
    """
    text = _SOURCES[name]
    eol = "\r\n" if "\r\n" in text else "\n"
    old = old.replace("\n", eol)
    new = new.replace("\n", eol)
    assert text.count(old) == count, (name, old[:80], text.count(old))
    return {name: text.replace(old, new)}


def test_00_the_accepted_sources_pass_every_detector() -> None:
    restore = _install()
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused == [], refused


# ===========================================================================
# 1-4. THE FOUR THINGS A RESET MUST NEVER TOUCH
# ===========================================================================
def test_01_reset_clears_a_business_input() -> None:
    """The register itself, wiped along with the results that came from it."""
    damaged = _swap(
        "modCalcReport",
        "    StateCell(CALC_STATE_ROW_LAST_ATTEMPT_RESULT).Value2 = CALC_ATTEMPT_NONE\n",
        "    StateCell(CALC_STATE_ROW_LAST_ATTEMPT_RESULT).Value2 = CALC_ATTEMPT_NONE\n"
        "    modWorkbook.Lo(SH_COST_LINES, TBL_COST_LINES).DataBodyRange.ClearContents\n")
    _control("test_13", sources=damaged)


def test_02_the_permanent_id_counter_resets() -> None:
    """Reissuing an identifier makes two different drivers the same driver."""
    damaged = _swap(
        "modReset",
        "    If ClearEveryPublication(calcUndo, simUndo, annualUndo, "
        "sensitivityUndo, detail) Then\n",
        "    modWorkbook.WriteValue nmCounterCostLine, 0\n"
        "    If ClearEveryPublication(calcUndo, simUndo, annualUndo, "
        "sensitivityUndo, detail) Then\n")
    _control("test_13", sources=damaged)


def test_03_the_auto_nonce_resets() -> None:
    """The nonce is monotonic because a discarded run must never be re-creatable."""
    damaged = _swap(
        "modSimReport",
        "    SharedCell(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT).Value2 = SIM_ATTEMPT_NONE\n",
        "    SharedCell(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT).Value2 = SIM_ATTEMPT_NONE\n"
        "    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = 0\n")
    _control("test_12", sources=damaged)


def test_04_the_run_id_counter_resets() -> None:
    """One row too high, and the whole anti-replay guarantee is gone."""
    damaged = _swap(
        "modSimReport",
        "        SIM_SHARED_VALUE_COLUMN & CStr(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT) & \":\" & _\n",
        "        SIM_SHARED_VALUE_COLUMN & CStr(SIM_IDENTITY_ROW_LAST_RUN_ID) & \":\" & _\n")
    _control("test_11", sources=damaged)


# ===========================================================================
# 5-6. ONE BANK IS NOT BOTH BANKS
# ===========================================================================
def test_05_only_the_active_simulation_bank_is_cleared() -> None:
    """The bank the selector names is the one the user can see. The other one is
    a complete published distribution that the next successful run would make
    visible again."""
    damaged = _swap(
        "modSimReport",
        "        SnapshotRange(SIM_BANK_A), SnapshotRange(SIM_BANK_B), _\n",
        "        SnapshotRange(SIM_BANK_A), _\n")
    _control("test_07", sources=damaged)


def test_06_the_inactive_bank_keeps_its_iteration_records() -> None:
    """Cleared identity, retained data: the worst of the two, because nothing on
    the sheet would say the rows were there."""
    damaged = _swap(
        "modSimReport",
        "        IterationRange(SIM_BANK_B, 0, SIM_MAX_ITERATIONS), _\n", "")
    _control("test_07", sources=damaged)


# ===========================================================================
# 7-9. A LAYER LEFT PUBLISHED
# ===========================================================================
def test_07_the_calculation_fingerprint_is_retained() -> None:
    """Retaining it makes the model read CURRENT against a calculation that is
    no longer on the sheet."""
    damaged = _swap(
        "modCalcReport",
        "        CALC_TOTALS_VALUE_RANGE, CALC_STATE_VALUE_RANGE)\n",
        "        CALC_TOTALS_VALUE_RANGE)\n")
    _control("test_05", sources=damaged)


def test_08_the_annual_publication_is_retained() -> None:
    damaged = _swap(
        "modSimAnnualStore",
        "        BlockAddress(SIM_BANK_A), BlockAddress(SIM_BANK_B))\n",
        "        StampAddress(SIM_BANK_B))\n")
    _control("test_08", sources=damaged)


def test_09_the_sensitivity_publication_is_retained() -> None:
    damaged = _swap(
        "modSimPostReport",
        "        RecordAddress(SIM_BANK_A), RecordAddress(SIM_BANK_B))\n",
        "        RecordAddress(SIM_BANK_A))\n")
    _control("test_09", sources=damaged)


# ===========================================================================
# 10-11. A SECOND OPINION ABOUT STATE, AND AN IDENTITY THAT MOVES
# ===========================================================================
def test_10_the_reset_forces_a_state_string() -> None:
    """The one failure that would make every state reading in the workbook a
    matter of who wrote last."""
    damaged = _swap(
        "modCalcReport",
        "    StateCell(CALC_STATE_ROW_LAST_ATTEMPT_RESULT).Value2 = CALC_ATTEMPT_NONE\n",
        "    StateCell(CALC_STATE_ROW_LAST_ATTEMPT_RESULT).Value2 = CALC_ATTEMPT_NONE\n"
        "    StateCell(CALC_STATE_ROW_CALCULATION_STATUS).Value2 = \"NOT CALCULATED\"\n")
    _control("test_16", sources=damaged)


def test_11_a_second_reset_advances_an_identity() -> None:
    """Pressing Reset twice must not be a way to consume a run id."""
    damaged = _swap(
        "modReset",
        "        ResetResults = modAppState.Succeeded(RESET_SUCCEEDED)\n",
        "        modSimReport.FinalCommit\n"
        "        ResetResults = modAppState.Succeeded(RESET_SUCCEEDED)\n")
    _control("test_18", sources=damaged)


# ===========================================================================
# 12-14. THE TRANSACTION
# ===========================================================================
def test_12_a_cancellation_still_mutates() -> None:
    """The confirmation moved below the first clear, which is exactly how a
    "Cancel" comes to mean "already done"."""
    damaged = _swap(
        "modReset",
        "    If Not modAppState.AskConfirm(ConfirmationSummary(), True) Then\n"
        "        ResetResults = modAppState.Succeeded(vbNullString)\n"
        "        Exit Function\n"
        "    End If\n\n"
        "    If ClearEveryPublication(calcUndo, simUndo, annualUndo, "
        "sensitivityUndo, detail) Then\n",
        "    If ClearEveryPublication(calcUndo, simUndo, annualUndo, "
        "sensitivityUndo, detail) Then\n"
        "        If Not modAppState.AskConfirm(ConfirmationSummary(), True) Then\n"
        "            ResetResults = modAppState.Succeeded(vbNullString)\n"
        "            Exit Function\n"
        "        End If\n")
    _control("test_24", sources=damaged)


def test_13_a_mid_reset_failure_leaves_a_partial_clear() -> None:
    """The rollback stops after the owner that failed, so everything cleared
    before it stays cleared."""
    damaged = _swap(
        "modReset",
        "    If Not modSimReport.SimReportRestorePublication(simUndo, note) Then _\n"
        "        problems = Appended(problems, note)\n"
        "    If Not modCalcReport.CalcReportRestorePublication(calcUndo, note) Then _\n"
        "        problems = Appended(problems, note)\n",
        "")
    _control("test_22", sources=damaged)


def test_14_the_rollback_restores_the_outputs_but_not_the_attempt_history() -> None:
    """A rolled-back reset that left "NONE" standing would have erased the record
    of a real earlier attempt while claiming nothing was lost."""
    damaged = _swap(
        "modCalcReport",
        "    For index = LBound(undo) To UBound(undo)\n"
        "        RestoredBlock undo(index)\n"
        "    Next index\n",
        "    For index = LBound(undo) To UBound(undo) - 1\n"
        "        RestoredBlock undo(index)\n"
        "    Next index\n")
    _control("test_42", sources=damaged)


# ===========================================================================
# 15-16. THE BUTTON, AND PROTECTION
# ===========================================================================
def test_15_the_button_is_bound_to_the_wrong_endpoint() -> None:
    """A destructive button on the wrong endpoint is the most expensive
    single-line mistake available in this contract."""
    damaged = _STRUCTURE.replace(
        '      entry_point: "PCCM_ResetResults"\n',
        '      entry_point: "PCCM_Calculate"\n')
    _control("test_30", structure=damaged)


def test_16_reset_unprotects_the_sheets_it_writes_to() -> None:
    """UserInterfaceOnly already lets every write below succeed. Unprotecting
    would leave the workbook open if the command then failed."""
    damaged = _swap(
        "modReset",
        "    If ClearEveryPublication(calcUndo, simUndo, annualUndo, "
        "sensitivityUndo, detail) Then\n",
        "    modProtection.ProtectionRelease detail\n"
        "    If ClearEveryPublication(calcUndo, simUndo, annualUndo, "
        "sensitivityUndo, detail) Then\n")
    _control("test_29", sources=damaged)


# ===========================================================================
# 17. THE PROJECTION IS AN AUTHORITY TOO
# ===========================================================================
def test_17_the_projection_quietly_drops_a_publication_group() -> None:
    """A Windows run can only look where the projection tells it to. A group
    missing from it is a store nobody would ever inspect."""
    damaged = json.loads(_PROJECTION)
    damaged["publications"]["simulation"]["cleared"].pop("sensitivity_records")
    _control("test_15", projection=json.dumps(damaged, indent=2) + "\n")


def test_18_the_projection_treats_a_recomputed_cell_as_a_preserved_input() -> None:
    """A run that compared a formula before and after a reset would fail on a
    workbook that was behaving correctly, and the real check would stop being
    believed."""
    damaged = json.loads(_PROJECTION)
    damaged["preserved"]["applied_timeline"]["cells"].append("C50")
    _control("test_14", projection=json.dumps(damaged, indent=2) + "\n")


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
