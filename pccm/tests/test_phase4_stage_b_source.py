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
        "LIMIT_MAX_YEAR_COLUMNS As Long = 301",
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
    assert "nextSequence = ReadCounter(Kind) + 1" in code


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
        "# C. Permanent Risk IDs",
        "# D - J. Timeline scenarios",
        "# K. Profiling synchronisation",
        "# L. Runtime failure containment",
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
    }
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        code = _ps_calls(path)
        for name in set(re.findall(r"\b([A-Z][a-z]+-[A-Z][\w]*)\b", code)):
            if name not in defined and name not in builtin:
                problems.append(f"{path.name}: {name} is invoked but never defined")
    assert not problems, "stale or undefined PowerShell helpers:\n  " + "\n  ".join(problems)


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
