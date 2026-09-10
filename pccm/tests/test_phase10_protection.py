#!/usr/bin/env python3
"""P10-2A: the protection foundation and the user command surface.

WHAT THIS BATCH ADDS, AND WHAT COULD GO WRONG WITH IT. Protection is the first
Phase-10 surface a user meets and the easiest to get subtly wrong: unlock one
formula and the model can be silently broken; lock one input and the workbook
cannot be used at all; forget UserInterfaceOnly and every accepted command stops
writing the moment the file is reopened. None of those failures announces
itself, so each is asserted here rather than assumed.

THE POLICY IS NOT RESTATED. These controls resolve the SAME declaration the
builder resolves and compare it against the workbook that was built from it. A
test carrying its own list of editable cells would be a third opinion, and the
first schema change would make one of the three wrong.

Runs standalone or under pytest.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
SRC = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"
BUILD = PCCM_ROOT / "build"
WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

from pccm_builder.contract_loader import load_contract  # noqa: E402
from pccm_builder.driver_loader import load_driver_contract  # noqa: E402
from pccm_builder.protection import resolve_unlocked  # noqa: E402
from pccm_builder.structure_loader import load_structure_contract  # noqa: E402
from openpyxl.utils import get_column_letter as _letter  # noqa: E402

WORKBOOK = BUILD / "PCCM_stageA.xlsx"
PROJECTION = BUILD / "phase10_protection_inspection.json"
MANIFEST = BUILD / "stage_b_manifest.json"

# THE FOUR COMMANDS THIS BATCH BINDS. Reset and Repair are NOT here: they are
# later steps, and a control that expected them now would fail for the right
# reason at the wrong time.
COMMANDS = (
    ("Calculate", "PCCM_Calculate", "btnPCCMCalculate"),
    ("Run Simulation", "PCCM_RunSimulation", "btnPCCMRunSimulation"),
    ("Run Sensitivity", "PCCM_RunSensitivity", "btnPCCMRunSensitivity"),
    ("Run Annual Cash Flow", "PCCM_RunAnnualStochastic", "btnPCCMRunAnnual"),
)

PHASE4_BUTTONS = (
    ("Setup", "btnPCCMApplyTimeline", "PCCM_ApplyTimeline", "Apply / Update Timeline", "E43"),
    ("Cost Lines", "btnPCCMAddCostLine", "PCCM_AddCostLine", "Add Cost Line", "N6"),
    ("Cost Lines", "btnPCCMDeleteCostLine", "PCCM_DeleteCostLine", "Delete Cost Line", "N9"),
    ("Risk Register", "btnPCCMAddRisk", "PCCM_AddRisk", "Add Risk", "O6"),
    ("Risk Register", "btnPCCMDeleteRisk", "PCCM_DeleteRisk", "Delete Risk", "O9"),
)


def _structure():
    return load_structure_contract(SPEC / "structure_contract.yaml")


def _resolved() -> dict[str, list[str]]:
    return resolve_unlocked(_structure(),
                            load_contract(SPEC / "input_contract.yaml"),
                            load_driver_contract(SPEC / "driver_contract.yaml"))


def _workbook():
    import openpyxl

    return openpyxl.load_workbook(WORKBOOK)


def _protection_source() -> str:
    return (SRC / "modProtection.bas").read_text(encoding="utf-8")


def _code_of(path: Path) -> str:
    """Executable VBA only.

    A BAN OR A REQUIREMENT MUST NOT BE SATISFIED BY THE COMMENT EXPLAINING IT.
    The first draft of test_08 asserted `UserInterfaceOnly:=True` against the
    raw file and passed with the flag deleted from the call, because the
    paragraph above it still said the words.
    """
    return "\n".join(line for line in path.read_text(encoding="utf-8").splitlines()
                      if not line.strip().startswith("'"))


def _event_source() -> str:
    return (SRC / "ThisWorkbook.vba").read_text(encoding="utf-8")


needs_build = pytest.mark.skipif(
    not WORKBOOK.is_file(), reason="Stage A has not been built into pccm/build")


# ===========================================================================
# A. EVERY INTENDED EDITABLE CELL IS UNLOCKED, BY DECLARATION
# ===========================================================================
@needs_build
def test_01_every_declared_editable_input_is_unlocked() -> None:
    """A. Resolved from the contracts, checked against the built file."""
    workbook = _workbook()
    wrong = []
    for sheet, addresses in _resolved().items():
        worksheet = workbook[sheet]
        for address in addresses:
            if worksheet[address].protection.locked:
                wrong.append(f"{sheet}!{address}")
    assert not wrong, f"declared editable cells are locked: {wrong[:8]}"


@needs_build
def test_01a_the_workbook_agrees_with_the_CONTRACTS_not_with_the_resolver() -> None:
    """THE INDEPENDENT HALF, AND THE FIRST DRAFT DID NOT HAVE IT.

    test_01 compares the resolver's answer with the workbook the resolver built.
    Both move together, so a resolver that unlocked the wrong cells satisfied
    it - two mutations survived exactly there. This asks the CONTRACTS instead:
    an input declared `editable: true` must be unlocked and one declared
    `editable: false` must be locked, whatever any resolver thinks.
    """
    contract = load_contract(SPEC / "input_contract.yaml")
    drivers = load_driver_contract(SPEC / "driver_contract.yaml")
    workbook = _workbook()
    wrong = []
    for spec in contract.inputs.values():
        locked = workbook[spec.sheet][spec.cell].protection.locked
        if locked == spec.editable:
            wrong.append(f"{spec.sheet}!{spec.cell} ({spec.key}) "
                         f"editable={spec.editable} locked={locked}")
    for sheet in ("Cost Lines", "Risk Register"):
        register = drivers.register_for_sheet(sheet)
        row = register.first_data_row
        for offset, column in enumerate(register.columns):
            address = f"{_letter(register.first_col_index + offset)}{row}"
            locked = workbook[sheet][address].protection.locked
            if locked == column.editable:
                wrong.append(f"{sheet}!{address} ({column.key}) "
                             f"editable={column.editable} locked={locked}")
    assert not wrong, f"the workbook disagrees with the contracts: {wrong[:8]}"


@needs_build
def test_02_the_editable_surface_is_not_empty_where_it_must_not_be() -> None:
    """AND THE RESOLUTION IS NOT VACUOUS. Every input sheet must resolve to
    something; a rule that quietly resolved to nothing would pass test_01 by
    unlocking nothing at all."""
    resolved = _resolved()
    for sheet in ("Setup", "Config", "Cost Lines", "Risk Register",
                  "Inflation", "Cost Profiling", "Risk Profiling"):
        assert resolved[sheet], f"{sheet} resolves to no editable cell"


# ===========================================================================
# B. WHAT MUST STAY LOCKED
# ===========================================================================
@needs_build
@pytest.mark.parametrize("sheet,address,what", [
    ("Setup", "C46", "an applied-timeline value the model controls"),
    ("Setup", "C50", "a derived formula"),
    ("Setup", "B9", "an input LABEL"),
    ("Cost Lines", "B12", "the permanent-id column"),
    ("Risk Register", "B12", "the permanent-id column"),
    ("Cost Profiling", "B13", "the permanent-id column"),
    ("Cost Profiling", "C13", "the description trace copy"),
    ("Inflation", "B13", "the synchronised profile name"),
    ("Model Check", "B18", "an output row"),
    ("Dashboard", "B8", "an output row"),
    ("Methodology", "B8", "a reference sheet"),
    ("_Calc", "B8", "a machine sheet"),
])
def test_03_representative_locked_cells_stay_locked(sheet: str, address: str,
                                                    what: str) -> None:
    """B. One probe per class of thing a user must not overwrite."""
    assert _workbook()[sheet][address].protection.locked, (
        f"{sheet}!{address} is unlocked, and it is {what}")


@needs_build
def test_04_no_output_reference_or_machine_sheet_unlocks_anything() -> None:
    resolved = _resolved()
    for sheet in ("Dashboard", "Model Check", "Results", "Sensitivity",
                  "Methodology", "_Calc", "_SimData"):
        assert resolved[sheet] == [], f"{sheet} unlocks {resolved[sheet][:4]}"


# ===========================================================================
# C-D. STRUCTURE, AND THE ABSENCE OF A PASSWORD
# ===========================================================================
def test_05_workbook_structure_protection_is_declared() -> None:
    policy = _structure().protection
    assert policy["protect_structure"] is True
    assert policy["user_interface_only"] is True
    assert policy["passwordless"] is True


def test_06_no_password_literal_exists_anywhere() -> None:
    """D. Not in the policy, not in the owner, not in the bootstrap, not in the
    event. A password in source would claim a security property this design
    explicitly does not have."""
    banned = re.compile(r"(?i)password\s*[:=]\s*['\"][^'\"]+['\"]")
    for path in (SPEC / "structure_contract.yaml", SRC / "modProtection.bas",
                 SRC / "ThisWorkbook.vba", WINDOWS / "build_stage_b.ps1"):
        text = path.read_text(encoding="utf-8")
        assert not banned.search(text), f"a password literal appears in {path.name}"
    # AND THE OWNER PASSES NO PASSWORD ARGUMENT AT ALL, empty or otherwise.
    source = _protection_source()
    assert "Password:=" not in source, "modProtection passes a password argument"


# ===========================================================================
# E. ONE OWNER
# ===========================================================================
def test_07_the_protection_owner_is_unique() -> None:
    """E and K. Protect/Unprotect appear in modProtection and nowhere else in
    the VBA. A second module spelling the policy is how two policies begin."""
    offenders = []
    for path in sorted(SRC.glob("*.bas")):
        if path.stem == "modProtection":
            continue
        body = "\n".join(line for line in path.read_text(encoding="utf-8").splitlines()
                         if not line.strip().startswith("'"))
        if re.search(r"\.(?:Protect|Unprotect)\b", body):
            offenders.append(path.name)
    assert not offenders, f"protection is spelled outside its owner: {offenders}"
    # AND THE OWNER HOLDS NO SHEET LIST AND NO RANGE. It applies one uniform
    # action; which cells are unlocked is already in the workbook.
    source = _protection_source()
    for spelled in ("Cost Lines", "Risk Register", "Cost Profiling", "Range("):
        assert spelled not in source, (
            f"modProtection spells {spelled!r}; the policy belongs in the spec")


def test_08_the_owner_applies_user_interface_only() -> None:
    """H. The flag every accepted command depends on after a reopen."""
    code = _code_of(SRC / "modProtection.bas")
    assert "UserInterfaceOnly:=True" in code, (
        "the flag is not in the Protect CALL; a comment mentioning it is not it")
    assert "ThisWorkbook.Protect Structure:=True" in code
    # AND IT IS ON THE SHEET PROTECT, not somewhere else in the module.
    protect_lines = [line for line in code.splitlines() if ".Protect " in line]
    sheet_protect = [line for line in protect_lines if line.strip().startswith("sheet.Protect")]
    assert sheet_protect, "no sheet is protected"
    assert all("UserInterfaceOnly:=True" in line for line in sheet_protect), sheet_protect
    source = _protection_source()
    for procedure in ("ProtectionApply", "ProtectionRelease", "ProtectionIsApplied"):
        assert f"Public Function {procedure}" in source, procedure


# ===========================================================================
# F-G. THE EVENT
# ===========================================================================
def test_09_the_event_delegates_and_does_nothing_else() -> None:
    """F and G. One job, and a named list of things it must not do."""
    source = _event_source()
    assert "Private Sub Workbook_Open()" in source
    assert source.count("Sub ") == 1, "ThisWorkbook grew a second procedure"
    assert "modProtection.ProtectionApply" in source
    body = "\n".join(line for line in source.splitlines()
                     if not line.strip().startswith("'"))
    for forbidden in ("PCCM_Calculate", "PCCM_RunSimulation", "PCCM_RunSensitivity",
                      "PCCM_RunAnnualStochastic", "PCCM_ResetResults",
                      "PCCM_RepairProfiling", "DeriveStatus", "Fingerprint",
                      "Publish", ".Calculate"):
        assert forbidden not in body, f"Workbook_Open reaches {forbidden}"
    # NO SUCCESS DIALOG. An operator opening a workbook asked for a workbook.
    assert body.count("ReportFailure") == 1
    assert "ReportResult" not in body, "the open handler reports a success"


def test_10_the_event_is_failure_safe_and_never_half_protects() -> None:
    source = _event_source()
    body = "\n".join(line for line in source.splitlines()
                     if not line.strip().startswith("'"))
    assert "ScreenUpdating = previousUpdating" in body, "application state is not restored"
    assert "ProtectionRelease" in body, (
        "a failed apply leaves the workbook half-protected")
    # AND IT FOLLOWS THE ACCEPTED AUTOMATION IDIOM rather than changing the
    # shared owner: the CALLER asks, exactly as modDrivers and modTimeline do,
    # so a harness never deadlocks on a modal dialog nobody can dismiss.
    assert "If Not modAppState.gAutomationActive Then" in body
    assert body.index("gAutomationActive") < body.index("ReportFailure")


def test_11_the_document_module_is_declared_and_is_not_a_bas() -> None:
    module = _structure().document_module
    assert module["component"] == "ThisWorkbook"
    assert module["file"] == "ThisWorkbook.vba"
    assert module["events"] == ["Workbook_Open"]
    assert module["delegates_to"] == "modProtection"
    assert (SRC / module["file"]).is_file()
    assert not (SRC / "ThisWorkbook.bas").exists(), (
        "the document module must not be in the .bas inventory")


# ===========================================================================
# I-J. THE BUTTONS
# ===========================================================================
@pytest.mark.parametrize("caption,entry,shape", COMMANDS)
def test_12_each_operational_command_is_declared_and_bound(caption: str, entry: str,
                                                           shape: str) -> None:
    """I. Declared, bound to the right existing endpoint, on Setup."""
    buttons = {b.shape_name: b for b in _structure().buttons}
    assert shape in buttons, f"{caption} has no button"
    button = buttons[shape]
    assert button.caption == caption
    assert button.entry_point == entry
    assert button.sheet == "Setup"
    # THE ENDPOINT REALLY EXISTS, and is Public. A button bound to a name that
    # is not there is a #NAME? the user meets by clicking.
    source = "\n".join(p.read_text(encoding="utf-8") for p in sorted(SRC.glob("*.bas")))
    assert re.search(rf"^Public (?:Sub|Function) {entry}\b", source, re.M), entry


def test_13_the_command_anchors_are_inside_the_declared_block() -> None:
    """I. Valid positions, and no collision with anything that exists."""
    structure = _structure()
    block = structure.commands["block"]
    first = int(block["first_button_row"])
    pitch = int(block["button_row_pitch"])
    column = str(block["button_column"])
    expected = [f"{column}{first + pitch * i}" for i in range(len(COMMANDS))]
    anchors = [b.anchor_cell for b in structure.buttons
               if b.shape_name in {s for _c, _e, s in COMMANDS}]
    assert anchors == expected, f"the command anchors drifted: {anchors}"
    # BELOW EVERYTHING THAT EXISTS. The applied-timeline block ends at row 54 and
    # the heading is at 56, so no declared cell is under a shape.
    assert int(block["section_row"]) > 54
    assert first > int(block["note_row"])
    # AND NOWHERE NEAR THE PHASE-4 BUTTON.
    others = {b.anchor_cell for b in structure.buttons
              if b.shape_name not in {s for _c, _e, s in COMMANDS}}
    assert not (set(expected) & others), "a command anchor collides with a Phase-4 button"


@pytest.mark.parametrize("sheet,shape,entry,caption,anchor", PHASE4_BUTTONS)
def test_14_every_phase4_button_is_unchanged(sheet: str, shape: str, entry: str,
                                             caption: str, anchor: str) -> None:
    """J. Byte for byte in the declaration: sheet, shape, caption, binding, cell."""
    buttons = {b.shape_name: b for b in _structure().buttons}
    assert shape in buttons, f"the Phase-4 button {shape} was removed"
    button = buttons[shape]
    assert (button.sheet, button.entry_point, button.caption, button.anchor_cell) == (
        sheet, entry, caption, anchor)


@needs_build
def test_15_the_command_block_is_rendered_where_the_buttons_anchor() -> None:
    block = _structure().commands["block"]
    worksheet = _workbook()["Setup"]
    label = str(block["label_column"])
    assert worksheet[f"{label}{block['section_row']}"].value == block["section"]
    assert worksheet[f"{label}{block['note_row']}"].value == block["note"]


# ===========================================================================
# L-M. WHAT THIS BATCH MAY NOT HAVE DISTURBED
# ===========================================================================
@needs_build
def test_16_the_phase9_read_only_adapters_are_untouched() -> None:
    """L. Protection changed no adapter, and no adapter learned to write."""
    adapter = (SRC / "modResultsState.bas").read_text(encoding="utf-8")
    for banned in (".Protect", ".Unprotect", ".Value2 =", "ThisWorkbook.Protect"):
        assert banned not in adapter, f"modResultsState now contains {banned}"
    assert adapter.count("Application.Volatile True") >= 8, (
        "an adapter lost its volatility")


def test_17_no_fingerprint_owner_was_touched() -> None:
    """M. Protection is presentation. If it had reached a fingerprint owner, a
    protected workbook and an unprotected one could disagree about identity."""
    for module in ("modCalcFingerprint", "modSimFingerprint", "modCalcReport",
                   "modSimReport"):
        body = (SRC / f"{module}.bas").read_text(encoding="utf-8")
        assert "Protect" not in body, f"{module} now mentions protection"


def test_18_no_command_unprotects_around_its_own_writes() -> None:
    """THE WORKAROUND THIS DESIGN EXISTS TO AVOID. UserInterfaceOnly is what
    lets every accepted command go on writing; a command that unprotected first
    would be evidence the flag is not being applied, and would leave the
    workbook unprotected if it failed midway."""
    for path in sorted(SRC.glob("*.bas")):
        if path.stem == "modProtection":
            continue
        body = "\n".join(line for line in path.read_text(encoding="utf-8").splitlines()
                         if not line.strip().startswith("'"))
        assert "Unprotect" not in body, f"{path.name} unprotects around its writes"


# ===========================================================================
# THE PROJECTION THE WINDOWS RUN WILL ASK
# ===========================================================================
@needs_build
def test_19_the_projection_matches_the_resolution_and_the_workbook() -> None:
    projection = json.loads(PROJECTION.read_text(encoding="utf-8"))
    resolved = _resolved()
    assert projection["passwordless"] is True
    assert projection["user_interface_only"] is True
    assert projection["protect_structure"] is True
    assert {s["sheet"] for s in projection["sheets"]} == set(resolved)
    for entry in projection["sheets"]:
        assert entry["protect"] is True
        assert entry["unlocked"] == resolved[entry["sheet"]]
        assert entry["unlocked_count"] == len(entry["unlocked"])


@needs_build
def test_20_the_manifest_carries_the_action_and_the_document_module() -> None:
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    protection = manifest["protection"]
    assert protection["passwordless"] is True
    assert protection["user_interface_only"] is True
    assert protection["protect_structure"] is True
    assert len(protection["sheets"]) == 14
    assert manifest["vba"]["document_module"]["component"] == "ThisWorkbook"
    assert {b["entry_point"] for b in manifest["buttons"]} >= {e for _c, e, _s in COMMANDS}
    # THE RESOLVED RANGES ARE NOT IN THE MANIFEST. They are in the workbook and
    # in the projection; a third copy is a third chance to disagree.
    assert "unlocked" not in json.dumps(protection)


def test_21_the_bootstrap_protects_last_and_reads_back() -> None:
    source = (WINDOWS / "build_stage_b.ps1").read_text(encoding="utf-8")
    assert "$manifest.protection" in source
    assert "UserInterfaceOnly" in source
    assert "[Type]::Missing" in source, "a null password would be marshalled as one"
    assert "will not invent one" in source, "the bootstrap does not refuse a password"
    # ORDER: protection after the buttons and before the save.
    assert source.index("Create the Phase-4 command buttons") < source.index(
        "Apply passwordless protection")
    assert source.index("Apply passwordless protection") < source.index("$wb.Save()")
    # AND THE DOCUMENT MODULE IS WRITTEN AND READ BACK, as the buttons are.
    assert "AddFromString" in source and "DeleteLines" in source
    assert "does not delegate to" in source


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
