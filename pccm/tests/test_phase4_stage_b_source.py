#!/usr/bin/env python3
"""PCCM Phase 4: static tests over the Stage-B source package.

VBA cannot be compiled or executed on Linux, and PowerShell cannot drive Excel
here. THESE TESTS DO NOT CLAIM THE STAGE-B RUNTIME IS CORRECT. Only a clean
Windows functional run can claim that.

What they do establish, mechanically:

  * the module inventory on disk is exactly the contract's, in both directions
  * every declared entry point exists, and no orphan PCCM_ macro exists
  * every SCREAMING_CASE constant the hand-written VBA references is actually
    emitted by the generated constants module -- the substitute for a compiler
    that would otherwise catch a typo only on Windows
  * no later-phase or forbidden construct has leaked into Phase-4 VBA
  * the bootstrap carries the proven COM lifecycle policy and alters no security
    setting
  * the functional harness covers every required scenario and hard-codes no
    expected timeline value

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import re
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    emit_stage_b,
    load_contract,
    load_driver_contract,
    load_spec,
    load_structure_contract,
)
from pccm_builder.stage_b_emit import build_manifest, render_constants_module  # noqa: E402
from pccm_builder.vba_source import contains_construct, load_modules  # noqa: E402

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"
CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
DRIVERS_PATH = PCCM_ROOT / "spec" / "driver_contract.yaml"
STRUCTURE_PATH = PCCM_ROOT / "spec" / "structure_contract.yaml"

SRC_VBA = PCCM_ROOT / "src" / "vba"
BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
LIFECYCLE_PS1 = BOOTSTRAP / "com_lifecycle.ps1"
BUILD_PS1 = BOOTSTRAP / "build_stage_b.ps1"
HARNESS_PS1 = BOOTSTRAP / "phase4_functional_test.ps1"

_EMITTED: dict[str, Path] = {}


def _specs():
    return (
        load_spec(SPEC_PATH),
        load_contract(CONTRACT_PATH),
        load_driver_contract(DRIVERS_PATH),
        load_structure_contract(STRUCTURE_PATH),
    )


def _emitted_dir() -> Path:
    """Emit the Stage-B artifacts into a scratch directory, once per run."""
    if "dir" in _EMITTED:
        return _EMITTED["dir"]
    tmp = Path(tempfile.mkdtemp(prefix="pccm-stageb-"))
    emit_stage_b(tmp, *_specs())
    _EMITTED["dir"] = tmp
    return tmp


def _generated_module_text() -> str:
    return (_emitted_dir() / "vba" / "modConstants.bas").read_text(encoding="utf-8")


def _manifest() -> dict:
    return json.loads((_emitted_dir() / "stage_b_manifest.json").read_text(encoding="utf-8"))


def _all_modules():
    return load_modules([SRC_VBA, _emitted_dir() / "vba"])


def _handwritten_modules():
    return load_modules([SRC_VBA])


def _ps(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _ps_code(path: Path) -> str:
    """PowerShell with block comments, line comments and trailing comments removed."""
    text = re.sub(r"<#.*?#>", "", _ps(path), flags=re.DOTALL)
    lines = []
    for line in text.splitlines():
        if line.lstrip().startswith("#"):
            continue
        lines.append(line.split(" #")[0])
    return "\n".join(lines)


def _ps_calls(path: Path) -> str:
    """PowerShell code with string literals emptied as well.

    Prose inside a quoted message ("Stage-B build", "Base-Year movement") matches
    the Verb-Noun shape of a cmdlet call, so an invocation sweep that did not strip
    strings would report those as undefined helpers.
    """
    code = _ps_code(path)
    code = re.sub(r"'(?:[^']|'')*'", "''", code)
    code = re.sub(r'"(?:[^"]|"")*"', '""', code)
    return code


# ===========================================================================
# module inventory
# ===========================================================================
def test_01_every_contract_module_exists_on_disk() -> None:
    structure = _specs()[3]
    available = {m.name for m in _all_modules()}
    declared = {m.name for m in structure.vba_modules}
    assert declared <= available, f"declared but missing: {sorted(declared - available)}"


def test_02_no_module_exists_that_the_contract_does_not_declare() -> None:
    structure = _specs()[3]
    declared = {m.name for m in structure.vba_modules}
    on_disk = {m.name for m in _all_modules()}
    assert on_disk <= declared, f"undeclared module(s) on disk: {sorted(on_disk - declared)}"


def test_03_exactly_one_module_is_generated_and_it_is_modconstants() -> None:
    structure = _specs()[3]
    generated = [m.name for m in structure.vba_modules if m.generated]
    assert generated == ["modConstants"]
    assert not (SRC_VBA / "modConstants.bas").exists(), (
        "modConstants is emitted from the contract; a hand-written copy would be a "
        "second definition of every structural literal"
    )


def test_04_the_generated_module_declares_itself_generated() -> None:
    text = _generated_module_text()
    assert "GENERATED FILE - DO NOT EDIT" in text
    assert "structure_contract.yaml" in text


def test_05_no_module_is_a_dumping_ground() -> None:
    """The split is by responsibility; one giant module would defeat the point."""
    structure = _specs()[3]
    assert len(structure.vba_modules) >= 6, "the responsibility split collapsed"
    for module in _handwritten_modules():
        lines = len(module.raw.splitlines())
        assert lines < 900, f"{module.name} is {lines} lines; split its responsibilities"


# ===========================================================================
# entry points
# ===========================================================================
def test_06_every_declared_entry_point_exists_in_the_source() -> None:
    structure = _specs()[3]
    available = {p for m in _all_modules() for p in m.public_procedures}
    for name in structure.entry_points:
        assert name in available, f"entry point {name} is not implemented"


def test_07_every_harness_procedure_exists_in_the_source() -> None:
    raw = STRUCTURE_PATH.read_text(encoding="utf-8")
    import yaml
    declared = yaml.safe_load(raw)["vba"]["harness_procedures"]
    available = {p for m in _all_modules() for p in m.public_procedures}
    for name in declared:
        assert name in available, f"harness procedure {name} is not implemented"


def test_08_no_orphan_pccm_macro_exists() -> None:
    """Every externally callable PCCM_ procedure is accounted for by the contract."""
    import yaml
    data = yaml.safe_load(STRUCTURE_PATH.read_text(encoding="utf-8"))
    accounted = set(data["vba"]["entry_points"]) | set(data["vba"]["harness_procedures"])
    found = {
        p for m in _all_modules() for p in m.public_procedures if p.startswith("PCCM_")
    }
    assert found <= accounted, f"undeclared PCCM_ macro(s): {sorted(found - accounted)}"


def test_09_each_entry_point_is_bound_to_exactly_one_button() -> None:
    structure = _specs()[3]
    bound = [b.entry_point for b in structure.buttons]
    assert sorted(bound) == sorted(structure.entry_points)
    assert len(set(bound)) == len(bound)


def test_10_the_five_required_buttons_are_declared_on_the_right_sheets() -> None:
    structure = _specs()[3]
    placement = {(b.sheet, b.caption) for b in structure.buttons}
    for required in (
        ("Setup", "Apply / Update Timeline"),
        ("Cost Lines", "Add Cost Line"),
        ("Cost Lines", "Delete Cost Line"),
        ("Risk Register", "Add Risk"),
        ("Risk Register", "Delete Risk"),
    ):
        assert required in placement, f"missing button {required}"
    assert len(structure.buttons) == 5, "no simulation button belongs in Phase 4"


# ===========================================================================
# the constants module is the only source of structural literals
# ===========================================================================
def test_11_every_constant_the_vba_references_is_emitted() -> None:
    """The substitute for a VBA compiler.

    A mistyped constant name would otherwise surface only as a Windows runtime
    error, after the review gate. Every SCREAMING_CASE identifier used in the
    hand-written modules must be emitted by modConstants or declared locally.
    """
    emitted = set(load_modules([_emitted_dir() / "vba"])[0].constants)
    # VBA and Excel names that are language or library members, not our constants.
    builtin = {
        "VBA", "MSG", "TRUE", "FALSE", "OK", "PCCM", "ID", "URL", "UI",
    }
    problems: list[str] = []
    for module in _handwritten_modules():
        local = set(module.constants)
        for name in sorted(module.referenced_upper_identifiers):
            if name in emitted or name in local or name in builtin:
                continue
            problems.append(f"{module.name}: {name}")
    assert not problems, (
        "these SCREAMING_CASE identifiers are referenced but never defined:\n  "
        + "\n  ".join(problems)
    )


def test_12_the_generated_module_is_deterministic() -> None:
    spec, contract, drivers, structure = _specs()
    first = render_constants_module(spec, contract, drivers, structure)
    second = render_constants_module(spec, contract, drivers, structure)
    assert first == second


def test_13_the_generated_module_carries_the_id_prefixes_and_limits() -> None:
    text = _generated_module_text()
    for fragment in (
        'ID_PREFIX_COST_LINE As String = "CL-"',
        'ID_PREFIX_RISK As String = "R-"',
        "LIMIT_MIN_YEAR As Long = 1900",
        "LIMIT_MAX_YEAR As Long = 2200",
        # Architecture Lock Revision B, independent of the calendar-year window.
        "LIMIT_MAX_YEAR_COLUMNS As Long = 200",
        "ID_COUNTER_MAX As Long = 2147483647",
    ):
        assert fragment in text, f"modConstants is missing: {fragment}"


def test_14_no_structural_literal_is_restated_in_hand_written_vba() -> None:
    """Sheet names, table names and defined names come from modConstants only."""
    structure = _specs()[3]
    literals = (
        {g.table_name for g in structure.all_grids}
        | {f.defined_name for f in structure.structural_fields}
        | {c.defined_name for c in structure.counters}
        | {"Cost Profiling", "Risk Profiling", "Risk Register", "_Calc"}
    )
    problems = []
    for module in _handwritten_modules():
        for literal in literals:
            if f'"{literal}"' in module.code_without_string_removal:
                problems.append(f"{module.name} hardcodes {literal!r}")
    assert not problems, "\n".join(problems)


# ===========================================================================
# scope discipline in the VBA
# ===========================================================================
def test_15_no_forbidden_construct_appears_in_phase_4_vba() -> None:
    structure = _specs()[3]
    modules = _all_modules()
    problems = []
    for construct in structure.forbidden_constructs:
        offenders = contains_construct(modules, construct)
        if offenders:
            problems.append(f"{construct} in {offenders}")
    assert not problems, "\n".join(problems)


def test_16_no_input_worksheet_change_automation_exists() -> None:
    """Structural operations are command-driven. There is no hidden side effect."""
    modules = _all_modules()
    for construct in ("Worksheet_Change", "Workbook_SheetChange", "Worksheet_SelectionChange"):
        assert not contains_construct(modules, construct), f"{construct} appears in code"


def test_17_no_calculation_or_simulation_code_leaked_in() -> None:
    modules = _all_modules()
    for construct in (
        "Rnd(", "Randomize", "MRG32k3a", "WorksheetFunction.Percentile",
        "ExpectedValue", "RunSimulation", "DiscountFactor", "EscalationFactor",
    ):
        assert not contains_construct(modules, construct), f"{construct} appears in code"


def test_18_the_vba_never_calls_finalreleasecomobject_or_changes_security() -> None:
    modules = _all_modules()
    for construct in ("FinalReleaseComObject", "AutomationSecurity", "EnableCancelKey"):
        assert not contains_construct(modules, construct)


def test_19_every_module_declares_option_explicit() -> None:
    for module in _all_modules():
        assert "Option Explicit" in module.raw, f"{module.name} lacks Option Explicit"


def test_20_the_structural_validator_reports_and_never_repairs() -> None:
    """A structural fault must fail the operation, not be silently corrected."""
    text = (SRC_VBA / "modStructuralCheck.bas").read_text(encoding="utf-8")
    assert "ValidateStructure" in text
    module = next(m for m in _handwritten_modules() if m.name == "modStructuralCheck")
    for repair in (".Value =", ".ClearContents", "ListRows.Add", "ListColumns.Add", ".Delete"):
        assert repair not in module.code, (
            f"modStructuralCheck writes to the workbook ({repair}); the validator "
            "reports faults and must never repair one"
        )


def test_21_every_structural_check_key_is_used_by_the_validator() -> None:
    structure = _specs()[3]
    module = next(m for m in _handwritten_modules() if m.name == "modStructuralCheck")
    emitted = load_modules([_emitted_dir() / "vba"])[0].constants
    for check in structure.structural_checks:
        constant = "CHK_" + check["key"].upper()
        assert constant in emitted, f"{constant} is not emitted"
        assert constant in module.code, f"{constant} is declared but never reported"


def test_22_the_apply_path_prevalidates_before_it_mutates() -> None:
    text = (SRC_VBA / "modTimeline.bas").read_text(encoding="utf-8")
    validate_at = text.index("PrevalidateEntered()")
    confirm_at = text.index("AskConfirm")
    write_at = text.index("WriteValue NM_APPLIED_BASE_YEAR")
    assert validate_at < confirm_at < write_at, (
        "the order must be prevalidate, then confirm, then modify: a cancellation "
        "needs no rollback only because nothing has moved when the user is asked"
    )


def test_23_delete_re_resolves_identity_after_confirmation() -> None:
    text = (SRC_VBA / "modDrivers.bas").read_text(encoding="utf-8")
    confirm_at = text.index("AskConfirm(summary")
    assert text.count("RowOfId(Kind, PermanentId)") >= 2, (
        "the row index must be resolved again after the prompt"
    )
    assert text.index("RowOfId(Kind, PermanentId)", confirm_at) > confirm_at


def test_24_the_counter_is_never_decremented_on_deletion() -> None:
    module = next(m for m in _handwritten_modules() if m.name == "modDrivers")
    code = module.code
    assert "counterBefore" in code, "the counter is restored only on a failed operation"
    for pattern in ("- 1", "-1"):
        assert f"WriteValue CounterName(Kind), ReadCounter(Kind) {pattern}" not in code
    assert "nextSequence = current + 1" in code


def test_24a_allocation_refuses_an_invalid_counter() -> None:
    """A corrupt counter must never silently become zero.

    CL-001 issued, CL-001 deleted, counter corrupted: current rows hold no ID, so a
    fallback of 0 plus a highest-issued of 0 would validate cleanly and the next Add
    would reissue CL-001. Current rows cannot testify about deleted history.
    """
    module = next(m for m in _handwritten_modules() if m.name == "modDrivers")
    code = module.code
    assert "TryReadCounter" in code
    assert "If Not TryReadCounter(Kind, current) Then" in code, (
        "AllocateId must refuse rather than assume a value"
    )
    assert "ReadLongInRange(CounterName(Kind), 0, ID_COUNTER_MAX, 0)" not in code, (
        "the silent zero fallback must be gone from the allocation path"
    )


def test_24b_allocation_refuses_cleanly_at_the_representation_ceiling() -> None:
    code = next(m for m in _handwritten_modules() if m.name == "modDrivers").code
    ceiling = code.index("If current >= ID_COUNTER_MAX Then")
    increment = code.index("nextSequence = current + 1")
    assert ceiling < increment, "the ceiling guard must precede counter + 1"


def test_24c_the_ceiling_is_described_as_a_representation_limit() -> None:
    text = _generated_module_text()
    assert "IMPLEMENTATION REPRESENTATION CEILING" in text
    assert "not a maximum on how many identifiers the model may issue" not in text, (
        "the old comment denied that the Long bound is a ceiling at all"
    )


def test_24d_structural_check_reports_an_invalid_counter_independently() -> None:
    code = next(m for m in _handwritten_modules() if m.name == "modStructuralCheck").code
    assert "CHK_COUNTER_INTEGRITY" in code
    assert "TryReadCounter" in code, (
        "the check must test the stored counter itself, not infer it from row count"
    )


def test_24e_a_malformed_id_tail_is_reported_not_ignored() -> None:
    code = next(m for m in _handwritten_modules() if m.name == "modDrivers").code
    assert "Unrepresentable" in code, (
        "an ID whose sequence cannot be represented must be surfaced, not skipped"
    )


def test_24f_no_counter_accessor_maps_invalid_state_to_a_value() -> None:
    """A lossy accessor is dangerous by existing, not only by being called.

    ReadCounter() returned 0 for a missing, blank or non-numeric counter. Allocation
    refused invalid state correctly, but the operation SNAPSHOT read through the lossy
    accessor -- so a rollback wrote 0 back over corrupt text and turned it into valid,
    exhausted-from-zero state. Any future caller of such an accessor would reopen the
    same hole, so the accessor itself is banned, not just that one call site.
    """
    lossy = re.compile(r"(?<!Try)\bReadCounter\s*\(")
    problems = []
    for module in _handwritten_modules():
        for number, line in enumerate(module.code.splitlines(), 1):
            if lossy.search(line):
                problems.append(f"{module.name}:{number}: {line.strip()[:70]}")
    assert not problems, (
        "a lossy counter accessor exists again:\n  " + "\n  ".join(problems)
    )


def test_24g_the_operation_snapshot_stores_the_counter_raw() -> None:
    """The snapshot must be byte-for-byte, so corruption comes back as corruption."""
    code = next(m for m in _handwritten_modules() if m.name == "modDrivers").code
    assert "Public Function RawCounter(ByVal Kind As String) As Variant" in code, (
        "the snapshot needs a Variant accessor that performs no conversion"
    )
    assert "Dim counterBefore As Variant" in code, (
        "typing the snapshot as Long would convert corrupt text into a valid number"
    )
    assert "counterBefore = RawCounter(Kind)" in code
    assert "Dim counterBefore As Long" not in code
    restore = code[code.index("Private Function TryRestoreDriver"):]
    assert "ByVal CounterBefore As Variant" in restore, (
        "the restore parameter must be Variant too, or the conversion just moves"
    )
    assert "modWorkbook.WriteValue CounterName(Kind), CounterBefore" in restore


def test_24h_a_counter_at_the_ceiling_is_not_a_structural_fault() -> None:
    """Exhausted is valid state. Reporting it as corruption would break every
    unrelated structural operation: revalidation runs after Apply and after Delete,
    and a fault there rolls the whole operation back."""
    code = next(m for m in _handwritten_modules() if m.name == "modStructuralCheck").code
    assert ">= ID_COUNTER_MAX" not in code and "= ID_COUNTER_MAX" not in code, (
        "the structural check must not treat the representation ceiling as a fault"
    )
    assert "CHK_COUNTER_INTEGRITY" in code, "an INVALID counter is still a fault"


def test_24i_the_runtime_never_evaluates_the_ceiling_plus_one() -> None:
    """ID_COUNTER_MAX + 1 overflows a VBA Long at runtime; the guard must precede it."""
    code = next(m for m in _handwritten_modules() if m.name == "modDrivers").code
    assert "If current >= ID_COUNTER_MAX Then" in code, (
        "a > test would let the ceiling value itself reach current + 1"
    )
    assert "If current > ID_COUNTER_MAX Then" not in code
    assert "ID_COUNTER_MAX + 1" not in code


def test_24j_no_comment_still_denies_the_ceiling() -> None:
    """The ceiling is real. Source that claims otherwise misleads the next reader."""
    for module in _handwritten_modules():
        assert "no artificial ID maximum anywhere" not in module.raw, module.name
    assert "no artificial ID maximum anywhere" not in _generated_module_text()
    generated = _generated_module_text()
    assert "A counter sitting exactly AT this value is VALID, exhausted state" in generated


def test_25_a_new_inflation_year_is_never_seeded_with_zero() -> None:
    module = next(m for m in _handwritten_modules() if m.name == "modInflation")
    assert "PROFILE_INITIAL_VALUE" not in module.code, (
        "the profiling initial value must not reach the inflation grid; a new "
        "escalation year is blank, never zero"
    )
    assert "slot.ClearContents" in module.code


def test_26_profiling_growth_seeds_zero_at_the_tail() -> None:
    module = next(m for m in _handwritten_modules() if m.name == "modProfiling")
    assert "PROFILE_INITIAL_VALUE" in module.code


# ===========================================================================
# application state, orphan data and error containment
# ===========================================================================
# `On Error Resume Next` is permitted ONLY in these procedures, each of which is a
# narrow existence probe or a non-mutating cosmetic step. Anywhere else it hides a
# failure the user needs to know about.
ON_ERROR_RESUME_NEXT_WHITELIST = {
    "modWorkbook": ["LoExists", "NameExists"],
    "modDrivers": ["AddDriver", "SelectedId"],
}


def test_26a_on_error_resume_next_appears_only_where_whitelisted() -> None:
    """The suppression that hid application-state restoration failures.

    RestoreAppState and RecalculateStructuralState both used it, so a failure to put
    Calculation, DisplayAlerts, EnableEvents, ScreenUpdating or StatusBar back was
    invisible and the operation still reported success.
    """
    procedure = re.compile(
        r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+\w+)\s+(\w+)",
        re.I,
    )
    problems = []
    for module in _handwritten_modules():
        current = "(module level)"
        for number, line in enumerate(module.code_without_string_removal.splitlines(), 1):
            match = procedure.match(line)
            if match:
                current = match.group(1)
            if "On Error Resume Next" in line:
                allowed = ON_ERROR_RESUME_NEXT_WHITELIST.get(module.name, [])
                if current not in allowed:
                    problems.append(f"{module.name}.{current} (line {number})")
    assert not problems, (
        "On Error Resume Next outside the documented whitelist:\n  " + "\n  ".join(problems)
    )


def test_26b_application_state_restoration_reports_every_failure() -> None:
    code = next(m for m in _handwritten_modules() if m.name == "modAppState").code
    assert "Public Function RestoreAppState" in code, "it must return a report, not be a Sub"
    for prop in ("Calculation", "DisplayAlerts", "EnableEvents", "ScreenUpdating", "StatusBar"):
        assert f"TryRestore{prop}" in code, f"{prop} has no individually reported restore"
    # All five are attempted; none short-circuits the rest.
    body = code[code.index("Public Function RestoreAppState"):code.index("Private Function TryRestoreCalculation")]
    assert body.count("failures = failures &") == 5


def test_26c_recalculation_failure_is_not_swallowed() -> None:
    code = next(m for m in _handwritten_modules() if m.name == "modAppState").code
    assert "Public Function RecalculateStructuralState() As String" in code
    assert "FinishOperation" in code, "one cleanup path both commands must route through"


def test_26d_a_failed_cleanup_makes_the_operation_fail() -> None:
    """A structural change that completed but was not cleaned up is not a success."""
    for name in ("modTimeline", "modDrivers"):
        code = next(m for m in _handwritten_modules() if m.name == name).code
        # The success path runs cleanup through the capture gate, and the gate's
        # report is what decides whether the operation reports success.
        assert "cleanup = FinishIfCaptured(snapshot, stateCaptured)" in code, name
        assert "If Len(cleanup) > 0 Then" in code, (
            f"{name}: a non-empty cleanup report must change the outcome"
        )
        assert "NOT left in a safe state" in next(
            m for m in _handwritten_modules() if m.name == name
        ).code_without_string_removal, name


def test_26e_cleanup_failures_never_hide_the_original_error() -> None:
    for name in ("modTimeline", "modDrivers"):
        text = next(m for m in _handwritten_modules() if m.name == name).code_without_string_removal
        assert "Cleanup ALSO reported problems" in text, name
        assert "restoreNote = restoreNote &" in text, (
            f"{name} must append the cleanup report, not replace the restore note"
        )


def test_26f_controlled_error_handling_starts_at_the_top_of_each_command() -> None:
    """Assessment reads user-controlled cells and can raise before any mutation."""
    timeline = next(m for m in _handwritten_modules() if m.name == "modTimeline").code
    handler = timeline.index("On Error GoTo AssessmentFailure")
    prevalidate = timeline.index("problems = PrevalidateEntered()")
    summary = timeline.index("summary = BuildSummary(")
    confirm = timeline.index("modAppState.AskConfirm(summary")
    mutation = timeline.index("On Error GoTo Failure")
    assert handler < prevalidate < summary < confirm < mutation, (
        "handling must be installed before prevalidation, assessment and confirmation"
    )
    drivers = next(m for m in _handwritten_modules() if m.name == "modDrivers").code
    assert drivers.index("On Error GoTo AssessmentFailure") < drivers.index("On Error GoTo Failure")


def test_26f1_the_handler_is_installed_before_application_state_is_captured() -> None:
    """CaptureAppState is itself fallible: it reads six Application properties.

    Installing the handler after the capture left a window where a failure escaped as
    a raw VBA runtime error -- the uncontrolled dialog the whole containment design
    exists to prevent. The true command boundary is the first statement of the
    command, before anything fallible.
    """
    for name in ("modTimeline", "modDrivers"):
        code = next(m for m in _handwritten_modules() if m.name == name).code
        handler = code.index("On Error GoTo AssessmentFailure")
        capture = code.index("snapshot = modAppState.CaptureAppState()")
        assert handler < capture, (
            f"{name}: the error handler must be installed before CaptureAppState"
        )


def test_26f2_cleanup_never_claims_a_snapshot_that_was_never_captured() -> None:
    """If capture failed there is no prior state, and saying it was restored is a lie.

    Calling the restore routine on an uninitialised snapshot would also write whatever
    the zero-valued struct happens to hold back onto Application.
    """
    for name in ("modTimeline", "modDrivers"):
        code = next(m for m in _handwritten_modules() if m.name == name).code
        assert "Dim stateCaptured As Boolean" in code, f"{name}: capture success is not tracked"
        assert "stateCaptured = True" in code
        assert "Private Function FinishIfCaptured(" in code, (
            f"{name}: cleanup must be gated on whether a snapshot exists"
        )
        # FinishOperation is reachable from exactly one place: inside the gate.
        calls = [
            line.strip()
            for line in code.splitlines()
            if "modAppState.FinishOperation(" in line
        ]
        assert calls == ["FinishIfCaptured = modAppState.FinishOperation(Snapshot)"], (
            f"{name}: FinishOperation must be called only from FinishIfCaptured, "
            f"found {calls}"
        )
        gate = code[code.index("Private Function FinishIfCaptured("):]
        gate = gate[: gate.index("End Function")]
        assert "If Not StateCaptured Then" in gate, (
            f"{name}: the gate must test the captured flag before restoring anything"
        )
        assert "never captured" in next(
            m for m in _handwritten_modules() if m.name == name
        ).code_without_string_removal, (
            f"{name}: the report must say a snapshot never existed, not imply one did"
        )
        for marker in ("cleanup = ", "assessCleanup = ", "failureCleanup = "):
            assert marker + "FinishIfCaptured(snapshot, stateCaptured)" in code, (
                f"{name}: a cleanup path bypasses the gate ({marker.strip()})"
            )


def test_26f3_delete_resolves_its_identity_inside_the_protected_shell() -> None:
    """`RunDriverOperation Kind, False, SelectedId(Kind)` is not equivalent.

    VBA evaluates the argument BEFORE entering the callee, so a failure to resolve the
    selection -- selection off the table, an error value in the ID cell -- escaped
    before any handler was installed. Resolution has to happen inside a shell that is
    already protected.
    """
    code = next(m for m in _handwritten_modules() if m.name == "modDrivers").code
    assert "Public Sub RunDeleteCommand(ByVal Kind As String)" in code
    shell = code[code.index("Public Sub RunDeleteCommand"):]
    shell = shell[: shell.index("End Sub")]
    assert shell.index("On Error GoTo ResolveFailed") < shell.index("SelectedId(Kind)"), (
        "the handler must cover the selection lookup"
    )
    offenders = [
        line.strip()
        for line in code.splitlines()
        if "SelectedId(" in line and "RunDriverOperation" in line
    ]
    assert not offenders, (
        "SelectedId must never be evaluated as a call argument:\n  " + "\n  ".join(offenders)
    )
    for button in ("PCCM_DeleteCostLine", "PCCM_DeleteRisk"):
        body = code[code.index(f"Public Sub {button}()"):]
        body = body[: body.index("End Sub")]
        assert "RunDeleteCommand" in body, f"{button} must go through the protected shell"
        assert "SelectedId" not in body, f"{button} resolves identity outside the shell"


def test_26f4_the_by_id_entry_points_bypass_selection_entirely() -> None:
    """The harness drives Delete by permanent ID, so a headless run never depends on
    a Windows selection state that automation does not set."""
    code = next(m for m in _handwritten_modules() if m.name == "modDrivers").code
    for entry in ("PCCM_DeleteCostLineById", "PCCM_DeleteRiskById"):
        body = code[code.index(f"Public Sub {entry}("):]
        body = body[: body.index("End Sub")]
        assert "SelectedId" not in body, f"{entry} must not consult the selection"
        assert "RunDriverOperation" in body


def test_26g_textof_is_error_safe() -> None:
    """Inspecting corruption must not crash on the bad cell it is inspecting."""
    code = next(m for m in _handwritten_modules() if m.name == "modWorkbook").code
    body = code[code.index("Public Function TextOf"):code.index("Public Function IsErrorText")]
    assert "IsError(Target.Value)" in body, "TextOf must handle an error value first"
    assert "ERROR_CELL_MARKER" in body, "an error cell needs a deterministic non-blank marker"
    assert "ERROR_CELL_MARKER" in _generated_module_text()


def test_26h_the_orphan_invariant_exists_and_is_contract_declared() -> None:
    structure = _specs()[3]
    keys = {c["key"] for c in structure.structural_checks}
    assert "no_orphan_structural_data" in keys
    assert "counter_integrity" in keys
    check = next(m for m in _handwritten_modules() if m.name == "modStructuralCheck").code
    assert "CHK_NO_ORPHAN_STRUCTURAL_DATA" in check
    workbook = next(m for m in _handwritten_modules() if m.name == "modWorkbook").code
    assert "Public Function OrphanRows" in workbook


def test_26i_every_mutating_command_runs_the_pre_mutation_gate() -> None:
    """Apply, Add and Delete all pass through the same targeted safety check."""
    for name in ("modTimeline", "modDrivers"):
        code = next(m for m in _handwritten_modules() if m.name == name).code
        assert "modStructuralCheck.PreMutationCheck()" in code, name
    check = next(m for m in _handwritten_modules() if m.name == "modStructuralCheck").code
    start = check.index("Public Function PreMutationCheck")
    gate = check[start:check.index("End Function", start)]
    assert "CheckOrphanRows()" in gate
    assert "ValidateStructure()" not in gate, (
        "a full validation here would block the legitimate 'Config profile removed, "
        "Apply will synchronise it' workflow"
    )


def test_26j_the_orphan_check_covers_all_five_structural_tables() -> None:
    check = next(m for m in _handwritten_modules() if m.name == "modStructuralCheck").code
    body = check[check.index("Private Function CheckOrphanRows"):check.index("Private Function OrphanFault")]
    for table in ("TBL_COST_LINES", "TBL_RISK_REGISTER", "TBL_COST_PROFILING",
                  "TBL_RISK_PROFILING", "TBL_INFLATION"):
        assert table in body, f"{table} is not covered by the orphan invariant"


def test_26k_add_refuses_when_an_orphan_exists_even_if_a_blank_row_follows() -> None:
    code = next(m for m in _handwritten_modules() if m.name == "modDrivers").code
    body = code[code.index("Private Function FirstFreeRow"):code.index("Public Function AddDriver")]
    assert "If OrphanRow > 0 Then Exit Function" in body, (
        "an orphan anywhere must block the add, not merely be recorded"
    )
    assert "firstBlank" in body, "the whole register must be scanned before a row is chosen"
    assert "If orphanRow > 0 Then" in code, (
        "AddDriver must fail on any orphan, not only when no free row was found"
    )


def test_26l_year_cells_get_the_input_treatment_explicitly() -> None:
    """Propagation is not relied on for a runtime-generated editable region."""
    workbook = next(m for m in _handwritten_modules() if m.name == "modWorkbook").code
    assert "Public Sub PaintYearCells" in workbook
    assert "FILL_INPUT" in workbook and "FILL_LOCKED" in workbook
    for name in ("modProfiling", "modInflation"):
        code = next(m for m in _handwritten_modules() if m.name == name).code
        assert code.count("modWorkbook.PaintYearCells") >= 2, (
            f"{name} must repaint after reshaping AND after row synchronisation"
        )
    text = _generated_module_text()
    assert "FILL_INPUT As Long" in text and "FILL_LOCKED As Long" in text


def test_26m_no_colour_literal_appears_in_vba_or_powershell() -> None:
    """The fills come from the presentation source authority, not from the code."""
    spec = _specs()[0]
    for colour in (spec.presentation["colors"]["input_fill"],
                   spec.presentation["colors"]["locked_fill"]):
        for module in _handwritten_modules():
            assert colour not in module.raw, f"{module.name} hardcodes #{colour}"
        for path in (BUILD_PS1, HARNESS_PS1):
            assert colour not in _ps_code(path), f"{path.name} hardcodes #{colour}"


def test_26n_the_snapshot_preserves_per_cell_presentation() -> None:
    code = next(m for m in _handwritten_modules() if m.name == "modWorkbook").code
    assert "Fills()" in code, "a restored row must come back with its input language"
    assert "s.Fills(r, c) = Target.DataBodyRange.Cells(r, c).Interior.Color" in code
    assert "Interior.Color = Snapshot.Fills(r, c)" in code


# ===========================================================================
# bootstrap and harness
# ===========================================================================
def test_27_both_scripts_share_one_com_lifecycle_implementation() -> None:
    for path in (BUILD_PS1, HARNESS_PS1):
        assert "com_lifecycle.ps1" in _ps(path), f"{path.name} does not dot-source the policy"
    assert LIFECYCLE_PS1.is_file()


def test_28_the_proven_lifecycle_policy_is_intact() -> None:
    code = _ps_code(LIFECYCLE_PS1)
    assert "Marshal]::ReleaseComObject" in code, "the only permitted release call is missing"
    assert "FinalReleaseComObject" not in code, "FinalReleaseComObject is prohibited"
    for name in ("Release-ComObjectSafe", "Release-Transient", "Invoke-NamedRelease",
                 "Wait-ExcelExit", "Test-IsOurExcelProcess", "Invoke-EmergencyExcelCleanup"):
        assert name in code, f"the lifecycle module lost {name}"


def test_29_no_generic_com_stack_was_reintroduced() -> None:
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        code = _ps_code(path)
        for banned in ("ComStack", "Push-ComObject", "Get-ComStackCount", "ReleasePlan",
                       "FinalReleaseComObject"):
            assert banned not in code, f"{path.name} reintroduces {banned}"


def test_30_every_powershell_script_releases_leaf_before_parent() -> None:
    for path in (BUILD_PS1, HARNESS_PS1):
        code = _ps_code(path)
        wb_close = code.index(".Close($false)")
        wb_release = code.index("'Workbook'") if "'Workbook'" in code else code.index("'Workbook2'")
        quit_at = code.index(".Quit()")
        app_release = code.index("'Application'")
        assert wb_close < wb_release, f"{path.name} releases the workbook before closing it"
        assert quit_at < app_release, f"{path.name} releases the application before Quit"


def test_31_no_script_alters_any_security_setting() -> None:
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        code = _ps_code(path)
        for banned in ("Set-ItemProperty", "New-ItemProperty", "Remove-ItemProperty",
                       "HKLM:", "HKCU:", "AccessVBOM", "VBAWarnings", "TrustedLocations",
                       "reg.exe", "regedit"):
            assert banned not in code, f"{path.name} touches {banned}"


def test_32_no_script_terminates_an_unidentified_excel_process() -> None:
    for path in (BUILD_PS1, HARNESS_PS1):
        code = _ps_code(path)
        assert "Stop-Process" not in code, (
            f"{path.name} stops a process directly; termination belongs to the "
            "identity-verified emergency path in com_lifecycle.ps1"
        )
    lifecycle = _ps_code(LIFECYCLE_PS1)
    stop_at = lifecycle.index("Stop-Process")
    guard_at = lifecycle.index("if (Test-IsOurExcelProcess $Identity)")
    assert guard_at < stop_at, "the identity guard must precede any force-stop"


def test_33_a_forced_stop_is_never_reported_as_success() -> None:
    lifecycle = _ps_code(LIFECYCLE_PS1)
    assert "This run is NOT a pass" in lifecycle
    for path in (BUILD_PS1, HARNESS_PS1):
        code = _ps_code(path)
        assert "EmergencyRequired" in code
        assert "NaturalExit" in code


def test_34_the_bootstrap_reads_everything_from_the_generated_manifest() -> None:
    code = _ps_code(BUILD_PS1)
    assert "stage_b_manifest.json" in code
    manifest = _manifest()
    for sheet in manifest["sheets"]:
        assert f'"{sheet["codename"]}"' not in code, f"{sheet['codename']} is hardcoded"
    for button in manifest["buttons"]:
        assert f'"{button["entry_point"]}"' not in code, f"{button['entry_point']} is hardcoded"


def test_35_the_bootstrap_performs_every_required_step() -> None:
    code = _ps_code(BUILD_PS1)
    for fragment in (
        "SaveAs", "xlsm_file_format", "_CodeName", "vbcomps.Import", "AddShape",
        "OnAction", "$wb.Save()", "$workbooks.Open(", "Wait-ExcelExit",
    ):
        assert fragment in code, f"the bootstrap is missing {fragment}"
    # It must reopen in a genuinely fresh instance, not reuse the build one.
    assert code.count("New-Object -ComObject Excel.Application") == 2


def test_36_the_bootstrap_reports_the_trust_center_prerequisite_without_changing_it() -> None:
    assert "Get-TrustAccessGuidance" in _ps_code(BUILD_PS1)
    guidance = _ps(LIFECYCLE_PS1)
    assert "Trust access to the VBA project object model" in guidance
    assert "will NOT change that setting" in guidance


def test_37_the_harness_covers_every_required_scenario() -> None:
    code = _ps(HARNESS_PS1)
    for marker in (
        "# A. Stage-B build",
        "# B. Permanent Cost Line IDs",
        "# B2. A REAL reorder of the Cost Lines table",
        "# C. Permanent Risk IDs",
        "# D - J. Timeline scenarios",
        "# K. Profiling synchronisation",
        "# L. Runtime failure containment",
        "# M. Real table growth beyond the reserved capacity",
        "# N. An Inflation Profile removed from Config is destructive",
        "# O. Non-numeric content in a removed profiling cell is a data loss",
        "# P. An oversized pasted timeline value is rejected cleanly",
        "# Q. Add failure after row mutation has begun",
        "# R. Delete failure after row mutation has begun",
        "# D0. Seed a REAL, KEYED Inflation Profile before the timeline scenarios",
        "# S. Application state is RESTORED, not forced to a convenient default",
        "# T. Unkeyed structural data blocks every mutating operation",
        "# U. A corrupt ID counter must never allow reuse",
        "# V. Generated year cells carry the EXACT editable-input treatment",
        "# W. The representation ceiling is EXHAUSTED VALID STATE, not corruption",
    ):
        assert marker in code, f"the harness is missing section: {marker}"


def test_38_the_harness_hardcodes_no_expected_timeline_value() -> None:
    """Expected years come from the oracle-derived fixture, never from the script."""
    code = _ps_code(HARNESS_PS1)
    years = sorted(set(re.findall(r"\b(?:19|20|21|22)\d{2}\b", code)))
    assert not years, (
        f"the harness contains calendar-year literals {years}; every expected value "
        "must come from build/phase4_scenarios.json"
    )
    assert "phase4_scenarios.json" in code


def test_39_the_harness_works_on_a_disposable_copy() -> None:
    code = _ps_code(HARNESS_PS1)
    assert "GetTempPath()" in code, "the harness must not modify the real build output"
    assert "Copy-Item" in code


def test_40_the_harness_checks_id_non_reuse_and_independence() -> None:
    code = _ps(HARNESS_PS1)
    assert "must NOT be reused" in code or "not reused" in code
    assert "independent of the cost sequence" in code


def test_41_the_harness_proves_cancellation_leaves_the_workbook_unchanged() -> None:
    code = _ps(HARNESS_PS1)
    assert "logically unchanged" in code
    assert "PERMANENTLY DELETED" in code, "it must assert the destructive prompt names the loss"


def test_42_the_harness_injects_a_failure_after_mutation_has_begun() -> None:
    code = _ps(HARNESS_PS1)
    assert "apply.after_profiling_columns" in code, (
        "the injected failure must land after the applied triple and the profiling "
        "columns have already changed, or it proves nothing about restore"
    )
    assert "logically restored" in code
    module = next(m for m in _handwritten_modules() if m.name == "modTimeline")
    assert "apply.after_profiling_columns" in module.code_without_string_removal


def test_43_the_injected_failure_hook_is_inert_in_normal_use() -> None:
    module = next(m for m in _handwritten_modules() if m.name == "modAppState")
    code = module.code
    assert "If Not gAutomationActive Then Exit Sub" in code, (
        "the fail point must return immediately unless automation was explicitly begun"
    )
    for other in _handwritten_modules():
        assert "gAutomationActive = True" not in other.code or other.name == "modAppState"


def test_44_the_harness_verifies_natural_shutdown() -> None:
    code = _ps_code(HARNESS_PS1)
    assert "Excel closed naturally" in _ps(HARNESS_PS1)
    assert "Wait-ExcelExit" in code


def test_44a_the_harness_exercises_a_real_listobject_reorder() -> None:
    """Editing a cell in place is not a reorder and proves nothing about identity."""
    code = _ps(HARNESS_PS1)
    assert "Invoke-TableSort" in code, "the harness must move whole rows"
    assert "$sortObj.Apply()" in _ps_code(HARNESS_PS1)
    assert "the physical row order actually changed" in code
    assert "each permanent ID still sits on its own row data" in code
    assert "FIRST ROW MARKER" not in code, (
        "the edit-in-place pseudo-reorder must be gone, not merely supplemented"
    )


def test_44b_the_harness_drives_the_register_past_its_reserved_capacity() -> None:
    """Until a 26th driver exists, the ListRows.Add path has never run."""
    code = _ps(HARNESS_PS1)
    assert "reserved_rows" in code
    assert "more identified rows than its reserved capacity" in code
    assert "the ListObject itself grew" in code
    assert "the grown row ID cell keeps the model-controlled treatment" in code
    assert "the grown row user cells keep the editable input treatment" in code
    assert "keeps its Data Validation on the grown row" in code
    assert "a generated profiling year cell equals the contract input_fill" in code


def test_44c_the_harness_covers_the_removed_config_profile_loss_path() -> None:
    code = _ps(HARNESS_PS1)
    assert "the confirmation names the removed profile" in code
    assert "cancelling leaves the inflation row and all its rates unchanged" in code
    assert "accepting removes the obsolete inflation row" in code


def test_44d_the_harness_covers_blank_preservation_and_non_numeric_loss() -> None:
    code = _ps(HARNESS_PS1)
    assert "an existing BLANK profiling cell is still blank after synchronisation" in code
    assert "text in a removed profiling cell triggers a destructive warning" in code


def test_44e_the_harness_covers_add_and_delete_failure_injection() -> None:
    code = _ps(HARNESS_PS1)
    for stage in ("add.after_write_id", "delete.after_remove"):
        assert stage in code, f"the harness never injects a failure at {stage}"
        assert any(stage in m.code_without_string_removal for m in _handwritten_modules()), (
            f"no VBA fail point is named {stage}"
        )
    assert "the driver table row count was restored" in code
    assert "the profiling row count was restored" in code
    assert "no identifier issued by the failed Add survives" in code


def test_44f_the_harness_covers_oversized_pasted_timeline_values() -> None:
    code = _ps(HARNESS_PS1)
    assert "rejected by prevalidation, not by an overflow" in code
    assert "an oversized Start Year is rejected cleanly" in code


def test_44g_the_harness_seeds_a_real_keyed_inflation_profile() -> None:
    """A rate written into a blank-profile row is not a keyed rate at all.

    Inflation ownership is (Profile Name, Calendar Year). Without a named profile
    in Config, CountRateLosses skips the row, SetYearColumns never captures it and
    SyncProfileRows clears it -- so the D-J scenarios proved nothing about
    calendar-year preservation and the destructive Base-Year step had no real rate
    to threaten.
    """
    code = _ps(HARNESS_PS1)
    assert "$testProfile" in code
    assert "Add-ConfigProfile" in code
    assert "the test profile is in the Config master" in code
    assert "the named inflation profile still has its row" in code


def test_44h_inflation_preservation_is_asserted_by_calendar_year() -> None:
    code = _ps(HARNESS_PS1)
    assert "$ratesBefore" in code, "a (Profile Name, Calendar Year) map must be captured"
    assert "every surviving calendar year keeps EXACTLY its own rate" in code
    assert "calendar years leaving the span are gone from the headers" in code
    assert "newly required inflation years arrive BLANK, never zero" in code


def test_44i_the_harness_proves_prior_application_state_is_restored() -> None:
    """Restoration of the caller's state, not coercion to a convenient default."""
    code = _ps(HARNESS_PS1)
    for prop in ("ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation", "StatusBar"):
        assert f"{prop} restored to its prior value" in code, prop
        assert f"{prop} restored after failure" in code, prop
    assert "PCCM harness sentinel" in code, "a non-default StatusBar proves restoration"
    assert "application state was restored (ScreenUpdating is on)" not in code, (
        "asserting a convenient default is not asserting restoration"
    )


def test_44j_the_harness_covers_all_three_orphan_classes() -> None:
    code = _ps(HARNESS_PS1)
    assert "Add is refused while a driver orphan exists" in code
    assert "Apply is refused while a profiling orphan exists" in code
    assert "Apply is refused while an inflation orphan exists" in code
    assert "once the orphans are cleared, Add succeeds again" in code


def test_44k_the_counter_scenario_covers_the_historical_case() -> None:
    """Deleting every identifier and then corrupting the counter is the danger."""
    code = _ps(HARNESS_PS1)
    assert "every identified Risk was deleted" in code
    assert "a valid counter with zero rows is not a fault" in code
    assert "Add Risk is refused while the counter is invalid" in code
    assert "a BLANK counter is refused too, never treated as zero" in code
    assert "does not reuse R-001" in code


def test_44k1_the_counter_scenario_proves_corruption_is_not_laundered() -> None:
    """The refusal is only half of it. The ROLLBACK is where reuse was reintroduced.

    A failed Add rolls the driver operation back, and the rollback restores the
    counter it snapshotted. Asserting only that Add was refused passed against source
    that then wrote a laundered 0 over the corrupt text. The counter itself has to be
    read back afterwards.
    """
    code = _ps(HARNESS_PS1)
    assert "the corrupt counter is STILL the same corrupt text after the failed Add" in code
    assert "-ceq 'corrupt'" in code, (
        "the readback must be case-sensitive and exact, not merely non-numeric"
    )
    assert "and structural revalidation still reports it as invalid" in code
    assert "the blank counter is STILL blank after the failed Add, not restored as 0" in code


def test_44n_the_harness_drives_the_representation_ceiling_at_runtime() -> None:
    """The ceiling has two halves that pull apart: refuse allocation, stay valid."""
    code = _ps(HARNESS_PS1)
    for check in (
        "the ceiling identifier is not already in the register",
        "a counter AT the ceiling is not reported as a structural fault",
        "Add is refused cleanly at the ceiling, with no overflow",
        "the refusal names the ceiling as a representation limit",
        "the counter is unchanged: nothing was allocated",
        "no register row was keyed",
        "the ceiling identifier was never issued",
        "no profiling row was created",
        "Apply Timeline still succeeds while the sequence is exhausted",
        "the counter was restored for the remaining scenarios",
    ):
        assert check in code, f"the ceiling scenario is missing: {check}"
    assert "$manifest.limits.id_counter_max" in _ps_code(HARNESS_PS1), (
        "the ceiling must come from the manifest, not be typed into the harness"
    )
    assert "2147483647" not in _ps_code(HARNESS_PS1)


def test_44l_year_cell_presentation_is_asserted_by_equality() -> None:
    code = _ps(HARNESS_PS1)
    assert "equals input_fill" in code
    assert "-ne $manifest.presentation.locked_fill" not in _ps_code(HARNESS_PS1), (
        "'not the locked fill' is not an assertion that it IS the input fill"
    )
    assert "a year cell on an UNKEYED reserved row is model-controlled" in code
    assert "tblInflation: a calendar-year cell on a NAMED profile row equals input_fill" in code


def test_44m_rollback_reasserts_the_phase3_input_contract() -> None:
    code = _ps(HARNESS_PS1)
    assert "Add-DriverRowContractChecks" in code
    assert "after Add rollback the restored row" in code
    assert "after Delete rollback the recreated row" in code
    assert "a restored profiling year cell keeps its number format" in code
    assert "a restored profiling year cell keeps a positive column width" in code
    assert "a restored profiling year cell on a KEYED row is editable-input styled" in code


# ===========================================================================
# mechanical sweeps
#
# Neither language can be compiled here, and the readiness phase produced two
# self-inflicted regressions of exactly this shape: a helper that was deleted
# while its call sites survived. Both sweeps are permanent for that reason.
# ===========================================================================
def test_45_every_qualified_vba_call_resolves_to_a_real_procedure() -> None:
    """modProfiling.SetYearColumns must exist in modProfiling, not just look right."""
    modules = {m.name: m for m in _all_modules()}
    problems = []
    for module in _handwritten_modules():
        for match in re.finditer(r"\b(mod[A-Z]\w*)\.(\w+)", module.code):
            target, member = match.group(1), match.group(2)
            if target not in modules:
                problems.append(f"{module.name}: calls into unknown module {target}")
                continue
            known = set(modules[target].procedures) | set(modules[target].constants)
            # Public module-level variables are legitimate targets too.
            known |= set(
                re.findall(r"^Public\s+(\w+)\s+As\s", modules[target].code, re.MULTILINE)
            )
            if member not in known:
                problems.append(f"{module.name}: {target}.{member} does not exist")
    assert not problems, "stale or mistyped cross-module references:\n  " + "\n  ".join(problems)


def test_46_every_powershell_helper_invoked_is_defined_somewhere() -> None:
    """A helper removed from com_lifecycle.ps1 must not leave live call sites."""
    defined: set[str] = set()
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        defined |= set(re.findall(r"^\s*function\s+([\w-]+)", _ps(path), re.MULTILINE))

    # Cmdlets and functions provided by PowerShell itself, used deliberately.
    builtin = {
        "Get-Content", "Set-Content", "Get-Process", "Stop-Process", "Get-Date",
        "Start-Sleep", "Write-Host", "Test-Path", "Remove-Item", "New-Item",
        "Copy-Item", "Join-Path", "Split-Path", "New-Object", "Add-Type",
        "Select-Object", "Where-Object", "ForEach-Object", "ConvertFrom-Json",
        "Set-StrictMode", "Get-CimInstance", "Out-Null", "Write-Verbose",
        "Sort-Object", "Measure-Object", "Select-String",
    }
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        code = _ps_calls(path)
        for name in set(re.findall(r"\b([A-Z][a-z]+-[A-Z][\w]*)\b", code)):
            if name not in defined and name not in builtin:
                problems.append(f"{path.name}: {name} is invoked but never defined")
    assert not problems, "stale or undefined PowerShell helpers:\n  " + "\n  ".join(problems)


# Variables that hold an Excel COM object. A chained member expression rooted at one
# of these mints an intermediate RCW that nothing names and nothing releases.
COM_ROOT_VARIABLES = (
    "excel", "excel2", "wb", "wb2", "workbooks", "workbooks2", "worksheets",
    "worksheets2", "localWorksheets", "ws", "ws2", "vbproj", "vbproj2", "vbcomps",
    "vbcomps2", "shapes", "shp", "anchor", "tf", "tr", "existing", "imported",
    "added", "body", "lo", "los", "cols", "rowsObj", "colsObj", "rng", "nm", "names",
    "cell", "interior", "validation", "sortObj", "sortFields", "keyRange", "hit",
    "target", "register", "Workbook", "ExcelApp", "c", "probe",
)


def _com_root_pattern() -> str:
    return "(" + "|".join(COM_ROOT_VARIABLES) + ")"


def test_45a_every_vba_block_construct_is_balanced() -> None:
    """A structural substitute for the compiler that cannot run here.

    An unclosed If, For, Do, With, Select Case, Type or procedure is a syntax error
    that would only surface on Windows, after the review gate. Line continuations
    are joined first so a wrapped condition is counted once.
    """
    openers = {
        "proc": re.compile(r"^(public |private |friend )?(static )?(sub|function) ", re.I),
        "type": re.compile(r"^(public |private )?type \w", re.I),
        "if": re.compile(r"^if .*\bthen$", re.I),
        "for": re.compile(r"^for\b", re.I),
        "do": re.compile(r"^do\b", re.I),
        "with": re.compile(r"^with\b", re.I),
        "select": re.compile(r"^select case\b", re.I),
    }
    closers = {
        "proc": re.compile(r"^end (sub|function)\b", re.I),
        "type": re.compile(r"^end type\b", re.I),
        "if": re.compile(r"^end if\b", re.I),
        "for": re.compile(r"^next\b", re.I),
        "do": re.compile(r"^loop\b", re.I),
        "with": re.compile(r"^end with\b", re.I),
        "select": re.compile(r"^end select\b", re.I),
    }
    problems = []
    for module in _all_modules():
        lines, buffer = [], ""
        for raw in module.code_without_string_removal.splitlines():
            stripped = raw.strip()
            if stripped.endswith("_"):
                buffer += stripped[:-1] + " "
                continue
            lines.append(buffer + stripped)
            buffer = ""
        if buffer:
            lines.append(buffer)

        counts = dict.fromkeys(openers, 0)
        for line in lines:
            for key, pattern in openers.items():
                if pattern.match(line):
                    counts[key] += 1
            for key, pattern in closers.items():
                if pattern.match(line):
                    counts[key] -= 1
        unbalanced = {k: v for k, v in counts.items() if v}
        if unbalanced:
            problems.append(f"{module.name}: {unbalanced}")
    assert not problems, "unbalanced VBA block constructs:\n  " + "\n  ".join(problems)


def test_45b_no_vba_line_exceeds_the_language_limit() -> None:
    """VBA rejects a physical line longer than 1023 characters."""
    problems = [
        f"{m.name}:{n}"
        for m in _all_modules()
        for n, line in enumerate(m.raw.splitlines(), 1)
        if len(line) > 1023
    ]
    assert not problems, f"lines beyond the VBA limit: {problems}"


def test_46a_no_chained_com_member_access_exists() -> None:
    """$Workbook.Names.Item(...) and $body.Rows.Count are the forbidden shape.

    Each creates an intermediate COM object that is never named, never released and
    never gated by a release failure. The proven readiness discipline requires
    acquire -> use -> release -> null on every one of them.
    """
    pattern = re.compile(rf"\${_com_root_pattern()}\.(\w+)\.(\w+)")
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        for number, line in enumerate(_ps_code(path).splitlines(), 1):
            for match in pattern.finditer(line):
                problems.append(
                    f"{path.name}:{number}: ${match.group(1)}.{match.group(2)}.{match.group(3)}"
                )
    assert not problems, "chained COM member access:\n  " + "\n  ".join(problems)


# Excel members that RETURN a COM object. Discarding the return leaks an RCW
# whether the call is bare, assigned to $null, or piped away.
OBJECT_RETURNING_MEMBERS = (
    "Import", "Add", "AddShape", "Open", "Item", "Cells", "Range", "Columns",
    "Rows", "ListRows", "ListColumns", "Offset", "Resize", "Find",
)


def test_46b_no_object_returning_com_call_is_left_unowned() -> None:
    """Discarding an object-returning call leaks an RCW, however it is discarded.

    The earlier sweep only matched a bare statement, so it missed
    `$null = $sortFields.Add(...)` -- SortFields.Add returns a SortField, and
    assigning it to $null mints the RCW and then throws ownership away.
    """
    members = "|".join(OBJECT_RETURNING_MEMBERS)
    discarded = re.compile(
        rf"(?:^|\|\s*|\$null\s*=\s*)\${_com_root_pattern()}\.({members})\("
    )
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        for number, line in enumerate(_ps_code(path).splitlines(), 1):
            stripped = line.strip()
            if discarded.search(stripped):
                problems.append(f"{path.name}:{number}: {stripped[:80]}")
    assert not problems, "unowned COM return value:\n  " + "\n  ".join(problems)


def test_46b1_the_known_object_returning_calls_are_captured_and_released() -> None:
    """Named regressions for the two that were actually found unowned."""
    build = _ps_code(BUILD_PS1)
    assert "$imported = $vbcomps.Import($file)" in build
    assert "Release-Transient $imported" in build

    harness = _ps_code(HARNESS_PS1)
    assert "$sortField = $sortFields.Add(" in harness, (
        "SortFields.Add returns a SortField and must be captured"
    )
    assert "Release-Transient $sortField" in harness
    assert "$null = $sortFields.Add(" not in harness, (
        "assigning an object-returning call to $null is not ownership"
    )


def test_46c_no_foreach_iterates_a_com_collection() -> None:
    """foreach over a COM collection hides the enumerator and every item RCW."""
    pattern = re.compile(rf"foreach\s*\(\s*\$\w+\s+in\s+\${_com_root_pattern()}\b")
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        for number, line in enumerate(_ps_code(path).splitlines(), 1):
            if pattern.search(line):
                problems.append(f"{path.name}:{number}: {line.strip()[:70]}")
    assert not problems, "foreach over a COM collection:\n  " + "\n  ".join(problems)


def test_46d_every_transient_release_nulls_the_caller_variable() -> None:
    """Release then null, on the same line, every time.

    PowerShell parameter binding cannot null a caller's variable, so a release that
    is not followed by an explicit assignment leaves a live alias to a released RCW.
    That is precisely the defect that produced the InvalidComObjectException in
    readiness run 2.
    """
    pattern = re.compile(r"Release-Transient\s+\$(\w+)\s+'[^']*'\s*;?\s*(.*)")
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        for number, line in enumerate(_ps_code(path).splitlines(), 1):
            match = pattern.search(line)
            if not match:
                continue
            variable, tail = match.group(1), match.group(2)
            if f"${variable}" not in tail or "$null" not in tail:
                problems.append(f"{path.name}:{number}: ${variable} released but not nulled")
    assert not problems, "released without nulling:\n  " + "\n  ".join(problems)


def test_46g_named_releases_also_null_their_variable() -> None:
    pattern = re.compile(r"Invoke-NamedRelease\s+\$\w+\s+\$(\w+)\s+'[^']*';\s*(.*)")
    problems = []
    for path in (BUILD_PS1, HARNESS_PS1):
        for number, line in enumerate(_ps_code(path).splitlines(), 1):
            match = pattern.search(line)
            if not match:
                continue
            variable, tail = match.group(1), match.group(2)
            if f"${variable}" not in tail or "$null" not in tail:
                problems.append(f"{path.name}:{number}: ${variable} released but not nulled")
    assert not problems, "released without nulling:\n  " + "\n  ".join(problems)


# Excel members that RETURN a COM object as a PROPERTY read rather than a call.
# Together with OBJECT_RETURNING_MEMBERS these are every acquisition shape in the
# two scripts, and each one mints an RCW the script then owns.
OBJECT_RETURNING_PROPERTIES = (
    "Worksheets", "Workbooks", "Shapes", "Names", "RefersToRange", "DataBodyRange",
    "HeaderRowRange", "ListObjects", "ListRows", "ListColumns", "VBProject",
    "VBComponents", "Properties", "Interior", "Validation", "Sort", "SortFields",
    "TextFrame2", "TextRange", "Font",
)


def _ps_structural_lines(path: Path) -> list[str]:
    """PowerShell lines with string CONTENTS blanked, preserving lines and columns.

    Brace tracking cannot run over raw source: a format string such as "{0} sheets"
    contributes braces that are not blocks. Blanking in place rather than deleting
    keeps every line number and column offset usable in a failure message.
    """
    out = []
    for line in _ps_code(path).splitlines():
        chars = list(line)
        quote = ""
        for index, char in enumerate(chars):
            if quote:
                chars[index] = " " if char != quote else char
                if char == quote:
                    quote = ""
            elif char in ("'", '"'):
                quote = char
        out.append("".join(chars))
    return out


def _ps_finally_marks(lines: list[str]) -> list[list[tuple[int, bool]]]:
    """Per line, the brace transitions and whether each lands inside a finally body.

    A release written inline in the normal flow is skipped when an earlier statement
    throws. A release written in a `finally` runs either way. Telling the two apart
    needs real brace tracking, including a `finally { ... }` written on one line.
    """
    marks: list[list[tuple[int, bool]]] = []
    stack: list[bool] = []
    pending = False
    token = re.compile(r"[{}]|(?<![\w-])finally(?![\w-])")
    for line in lines:
        per: list[tuple[int, bool]] = []
        for match in token.finditer(line):
            found = match.group(0)
            if found == "finally":
                pending = True
                continue
            if found == "{":
                stack.append(pending)
                pending = False
            elif stack:
                stack.pop()
            per.append((match.start(), any(stack)))
        marks.append(per)
    return marks


def _inside_finally(marks: list[list[tuple[int, bool]]], index: int, column: int) -> bool:
    state = False
    for earlier in marks[:index]:
        if earlier:
            state = earlier[-1][1]
    for position, after in marks[index]:
        if position < column:
            state = after
        else:
            break
    return state


def _ps_release_map(path: Path) -> tuple[set[str], dict[str, int]]:
    """(variables released in a finally, variables released only inline)."""
    lines = _ps_structural_lines(path)
    marks = _ps_finally_marks(lines)
    in_finally: set[str] = set()
    inline: dict[str, int] = {}
    pattern = re.compile(r"(?<![\w-])Release-Transient\s+\$(\w+)")
    for index, line in enumerate(lines):
        for match in pattern.finditer(line):
            variable = match.group(1)
            if _inside_finally(marks, index, match.start()):
                in_finally.add(variable)
            else:
                inline.setdefault(variable, index + 1)
    return in_finally, inline


def test_46h_no_com_release_sits_only_on_the_success_path() -> None:
    """Every inline release needs a guarded backstop in the enclosing finally.

    An inline release is a legitimate early hand-back only when the enclosing finally
    also releases the same variable if it is still non-null. Without that backstop the
    RCW leaks the moment anything between acquisition and release throws -- which is
    exactly how the pre-existing VBComponent and Shape leaked.
    """
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        in_finally, inline = _ps_release_map(path)
        for variable, number in sorted(inline.items()):
            if variable not in in_finally:
                problems.append(
                    f"{path.name}:{number}: ${variable} is released inline with no "
                    "release in the enclosing finally"
                )
    assert not problems, "COM release on the success path only:\n  " + "\n  ".join(problems)


def test_46i_every_com_acquisition_reaches_an_exception_safe_release() -> None:
    """Ownership starts at the assignment, not at the first successful use."""
    members = "|".join(OBJECT_RETURNING_MEMBERS + OBJECT_RETURNING_PROPERTIES)
    acquisition = re.compile(
        rf"^\s*\$(\w+)\s*=\s*\${_com_root_pattern()}\.({members})\b"
    )
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        in_finally, _ = _ps_release_map(path)
        ledger = set(
            re.findall(r"(?<![\w-])Invoke-NamedRelease\s+\$\w+\s+\$(\w+)", _ps_code(path))
        )
        raw = _ps_code(path).splitlines()
        for number, line in enumerate(_ps_structural_lines(path), 1):
            match = acquisition.match(line)
            if not match:
                continue
            variable = match.group(1)
            if variable in in_finally or variable in ledger:
                continue
            problems.append(f"{path.name}:{number}: {raw[number - 1].strip()[:70]}")
    assert not problems, (
        "COM object acquired with no exception-safe release:\n  " + "\n  ".join(problems)
    )


def test_46k_every_powershell_block_is_balanced() -> None:
    """The Windows scripts cannot be parsed here, so blocks are counted instead.

    An unclosed try, finally, function or foreach is a parse error that would only
    surface on Windows, after this review gate -- the same reason the VBA block
    sweep exists.
    """
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        depth = 0
        for number, line in enumerate(_ps_structural_lines(path), 1):
            for char in line:
                if char == "{":
                    depth += 1
                elif char == "}":
                    depth -= 1
                    assert depth >= 0, f"{path.name}:{number}: closes a block that never opened"
        assert depth == 0, f"{path.name}: {depth} block(s) left open"


def test_46j_the_two_exception_path_leaks_are_closed_by_name() -> None:
    """Named regressions for the two defects the sweep above was written from."""
    build = _ps_code(BUILD_PS1)

    # A. the pre-existing VBComponent removed before a module is re-imported.
    replace = build[build.index("$existing = $vbcomps.Item($moduleName)") :]
    replace = replace[: replace.index("$imported = $vbcomps.Import($file)")]
    assert "$vbcomps.Remove($existing)" in replace
    assert "} finally {" in replace, (
        "Remove() can throw; the acquire/use path needs a finally"
    )
    remove_at = replace.index("$vbcomps.Remove($existing)")
    finally_at = replace.index("} finally {")
    assert remove_at < finally_at, "the finally must follow the use, not precede it"
    assert "Release-Transient $existing 'VBComponent(existing)'; $existing = $null" in (
        replace[finally_at:]
    ), "the VBComponent must be released from the finally"

    # B. the pre-existing Shape deleted before a button is recreated.
    button = build[build.index("$existing = $shapes.Item($button.shape_name)") :]
    button = button[: button.index("Add-Step 'Create the Phase-4 command buttons'")]
    delete_at = button.index("$existing.Delete()")
    tail = button[delete_at:]
    assert "Release-Transient $existing" not in tail[: tail.index("} finally {")], (
        "an inline release after Delete() is skipped when Delete() throws"
    )
    cleanup = button[button.index("} finally {") :]
    assert "Release-Transient $existing 'Shape(existing)'" in cleanup
    assert cleanup.index("Release-Transient $existing") < cleanup.index(
        "Release-Transient $shapes"
    ), "leaf before parent: the Shape must release before the Shapes collection"


def test_46e_the_generated_module_resolves_from_the_supplied_build_dir() -> None:
    """The disposable harness build must test its OWN generated modConstants.bas.

    Resolving the generated directory against the repository root meant the harness
    copied a build to %TEMP% and then imported modConstants.bas from the real
    repository build -- a different generated source than the manifest and scenario
    fixture sitting beside the workbook it was driving.
    """
    code = _ps_code(BUILD_PS1)
    assert "$genDir  = Join-Path $BuildDir" in code, (
        "the generated VBA directory must resolve from the supplied BuildDir"
    )
    assert "$genDir  = Join-Path $pccmRoot" not in code
    # Source modules stay repository-relative: they are version-controlled input.
    assert "$srcDir  = Join-Path $pccmRoot $manifest.vba.source_dir" in code


def test_46f_the_harness_copies_a_coherent_build_set() -> None:
    code = _ps_code(HARNESS_PS1)
    for item in ("stage_a_filename", "stage_b_manifest.json", "phase4_scenarios.json", "'vba'"):
        assert item in code, f"the disposable copy is missing {item}"


def test_47_no_powershell_helper_is_defined_but_never_used() -> None:
    """A helper with no call sites is either dead or a symptom of a deleted one."""
    text = "\n".join(_ps_calls(p) for p in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1))
    defined = set(re.findall(r"^\s*function\s+([\w-]+)", text, re.MULTILINE))
    unused = [name for name in defined if len(re.findall(rf"\b{re.escape(name)}\b", text)) < 2]
    assert not unused, f"defined but never invoked: {sorted(unused)}"


# ===========================================================================
# emitted artifacts
# ===========================================================================
def test_48_the_manifest_matches_the_contracts() -> None:
    spec, contract, drivers, structure = _specs()
    expected = build_manifest(spec, contract, drivers, structure)
    assert _manifest() == json.loads(json.dumps(expected))


def test_48a_the_manifest_publishes_the_representation_ceiling() -> None:
    """One source for the ceiling: the VBA constant and the harness read the same
    number, so a change to it cannot leave the functional test asserting the old one."""
    from pccm_builder.stage_b_emit import VBA_LONG_MAX

    limits = _manifest()["limits"]
    assert limits["id_counter_max"] == VBA_LONG_MAX
    assert f"Public Const ID_COUNTER_MAX As Long = {VBA_LONG_MAX}" in _generated_module_text()
    # It is a limit of the implementation, and the other two limits are unrelated
    # to it: the calendar window and the generation guard stay exactly as locked.
    assert limits["min_year"] == 1900
    assert limits["max_year"] == 2200
    assert limits["max_generated_year_columns"] == 200


def test_49_the_manifest_declares_the_macro_enabled_file_format() -> None:
    assert _manifest()["xlsm_file_format"] == 52
    assert _manifest()["stage_b_filename"].endswith(".xlsm")


def test_50_the_manifest_carries_all_fourteen_codenames() -> None:
    manifest = _manifest()
    assert len(manifest["sheets"]) == 14
    for sheet in manifest["sheets"]:
        assert re.match(r"^sh[A-Z]", sheet["codename"]), sheet


def test_51_the_manifest_exposes_both_entered_aliases_and_applied_names() -> None:
    names = _manifest()["defined_names"]
    for name in ("nmBaseYear_Entered", "nmStartYear_Entered", "nmDuration_Entered",
                 "nmBaseYear_Applied", "nmStartYear_Applied", "nmDuration_Applied",
                 "nmStructuralState", "nmCounterCostLine", "nmCounterRisk"):
        assert name in names, f"{name} is missing from the manifest"


def test_52_the_stage_a_build_emits_all_three_artifacts() -> None:
    directory = _emitted_dir()
    for relative in ("vba/modConstants.bas", "stage_b_manifest.json", "phase4_scenarios.json"):
        assert (directory / relative).is_file(), f"{relative} was not emitted"


def test_53_the_generated_module_is_valid_vba_shape() -> None:
    text = _generated_module_text()
    assert text.startswith('Attribute VB_Name = "modConstants"')
    assert "\nOption Explicit\n" in text
    # Every non-comment, non-blank line must be a declaration.
    for line in text.splitlines()[2:]:
        stripped = line.strip()
        if not stripped or stripped.startswith("'"):
            continue
        assert stripped.startswith("Public Const"), f"unexpected line: {line}"


def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 4 static tests - Stage-B source package")
    print("=" * 70)
    for name, fn in tests:
        try:
            fn()
        except AssertionError as error:
            failures += 1
            print(f"  [FAIL] {name}\n         {error}")
        except Exception as error:  # noqa: BLE001
            failures += 1
            print(f"  [ERROR] {name}\n          {type(error).__name__}: {error}")
        else:
            print(f"  [PASS] {name}")
    print("=" * 70)
    print(f"  {len(tests) - failures} passed, {failures} failed")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(_run_all())
