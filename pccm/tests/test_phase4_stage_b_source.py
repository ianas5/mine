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
from pccm_builder.calc_emit import emit_calc_artifacts  # noqa: E402
from pccm_builder.sim_emit import emit_sim_artifacts  # noqa: E402
from pccm_builder.sim_loader import load_sim_contract  # noqa: E402
from pccm_builder.structure_loader import GENERATED_MODULES  # noqa: E402
from pccm_builder.calc_loader import load_calc_contract  # noqa: E402
from pccm_builder.stage_b_emit import build_manifest, render_constants_module  # noqa: E402
from pccm_builder.vba_source import (  # noqa: E402
    contains_construct,
    load_modules,
    logical_statements,
)

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"
CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
DRIVERS_PATH = PCCM_ROOT / "spec" / "driver_contract.yaml"
STRUCTURE_PATH = PCCM_ROOT / "spec" / "structure_contract.yaml"
CALC_PATH = PCCM_ROOT / "spec" / "calc_contract.yaml"
SIM_PATH = PCCM_ROOT / "spec" / "sim_contract.yaml"

SRC_VBA = PCCM_ROOT / "src" / "vba"
BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
LIFECYCLE_PS1 = BOOTSTRAP / "com_lifecycle.ps1"
BUILD_PS1 = BOOTSTRAP / "build_stage_b.ps1"
HARNESS_PS1 = BOOTSTRAP / "phase4_functional_test.ps1"
# The Phase-5 Gate-B scenarios, DOT-SOURCED into the harness above. It is not a
# second harness: it shares that script's scope, its helpers, its one COM
# lifecycle and its reporting, so a helper it defines is defined for the harness
# and a helper it calls must exist somewhere in this set.
SCENARIOS_PS1 = BOOTSTRAP / "phase5_gate_b_scenarios.ps1"

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
    spec, contract, drivers, structure = _specs()
    emit_stage_b(tmp, spec, contract, drivers, structure)
    # The Stage-B build emits TWO generated modules. `emit_stage_b` owns
    # modConstants and `emit_calc_artifacts` owns modCalcContract; the second is
    # called here rather than folded into the first so there is still exactly one
    # generator per artifact and no chance of a duplicate modCalcContract.
    emit_calc_artifacts(tmp, spec, load_calc_contract(CALC_PATH))
    # Phase 6 Step 5 emits a THIRD generated module. It entered the Stage-B
    # registry in Step 6, when modSimRng - the first module that depends on it -
    # arrived, so the inventory and constant-reference checks below must see it.
    emit_sim_artifacts(
        tmp, spec, load_sim_contract(SIM_PATH), contract, load_calc_contract(CALC_PATH)
    )
    _EMITTED["dir"] = tmp
    return tmp


def _generated_module_text() -> str:
    return (_emitted_dir() / "vba" / "modConstants.bas").read_text(encoding="utf-8")


def _manifest() -> dict:
    return json.loads((_emitted_dir() / "stage_b_manifest.json").read_text(encoding="utf-8"))


def _generated_modules():
    """Every module the Stage-A build emits, by name.

    Indexed by name rather than by position: there is more than one generated
    module now, and `[0]` silently became modCalcContract when it was added.
    """
    return {m.name: m for m in load_modules([_emitted_dir() / "vba"])}


def _public_constants(module) -> set[str]:
    """The Public Const names a module exports.

    A `Public Const` in one standard module is visible in every other, exactly as
    a compiler would see it; a `Private Const` is not, and stays invisible here
    too. Phase 4 never needed the distinction because its hand-written modules
    referenced only modConstants.
    """
    return {
        match.group(1)
        for line in module.code_without_string_removal.splitlines()
        if (match := re.match(r"^\s*Public\s+Const\s+(\w+)", line, re.IGNORECASE))
    }


def _generated_constants() -> set[str]:
    """Every constant projected by any generated module."""
    return {name for module in _generated_modules().values() for name in module.constants}


def _all_modules():
    return load_modules([SRC_VBA, _emitted_dir() / "vba"])


def _handwritten_modules():
    return load_modules([SRC_VBA])


def _ps(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _ps_code(path: Path) -> str:
    """PowerShell with block comments, line comments and trailing comments removed.

    LINE NUMBERING IS PRESERVED. Comments are blanked in place rather than deleted,
    because every sweep built on this reports a line number, and a sweep that sends
    the reader to the wrong line is worth much less than one that does not. The
    block comment is replaced by its own newlines for the same reason.
    """
    text = re.sub(
        r"<#.*?#>", lambda m: "\n" * m.group(0).count("\n"), _ps(path), flags=re.DOTALL
    )
    lines = []
    for line in text.splitlines():
        if line.lstrip().startswith("#"):
            lines.append("")
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


def test_03_the_generated_modules_are_exactly_the_ones_the_builder_emits() -> None:
    """The DEPLOYMENT invariant, carried forward from "exactly one generated module".

    That earlier assertion said two things at once: modConstants must be generated,
    and nothing else may claim to be. Phase 5 emits a second generated module, so
    the first half survives verbatim and the second half becomes a comparison
    against the builder's own locked list instead of against the number one.
    A hand-written copy of either is still refused.
    """
    structure = _specs()[3]
    generated = [m.name for m in structure.vba_modules if m.generated]
    assert structure.vba_generated_module == "modConstants"
    assert "modConstants" in generated
    assert sorted(generated) == sorted(GENERATED_MODULES)
    for name in generated:
        assert not (SRC_VBA / f"{name}.bas").exists(), (
            f"{name} is emitted from a contract; a hand-written copy would be a "
            "second definition of every literal it projects"
        )


def test_04_the_generated_module_declares_itself_generated() -> None:
    text = _generated_module_text()
    assert "GENERATED FILE - DO NOT EDIT" in text
    assert "structure_contract.yaml" in text


# The Phase-4 modules, whose limit is NOT relaxed by anything below.
PHASE4_VBA_MODULES = (
    "modWorkbook", "modAppState", "modTimeline", "modDrivers",
    "modProfiling", "modInflation", "modStructuralCheck",
)

# The Phase-5 modules. Two limits, both enforced. The resolver is measured by
# the same policy as the numerical kernel: it is a Phase-5 module whose
# responsibility is coherent and whose contract requires it to explain itself.
PHASE5_VBA_MODULES = (
    "modCalcFactors", "modCalcAnalytical", "modCalcFingerprint", "modCalcResolve",
    "modCalcCheck", "modCalcReport",
)

# The Phase-6 modules. Measured by the same policy as the Phase-5 kernel: a
# coherent responsibility whose contract requires it to explain itself at
# length. modSimRng carries the orientation, reduction and scope reasoning a
# later reader cannot reconstruct from the code alone.
PHASE6_VBA_MODULES = (
    "modSimRng",
    "modSimSample",
    "modSimEngine",
    "modSimStats",
)

PHASE4_RAW_LINE_LIMIT = 900
PHASE5_CODE_LINE_LIMIT = 900
PHASE5_RAW_LINE_LIMIT = 1200


def _line_metrics(module) -> tuple[int, int, int, int]:
    """(raw, blank, comment, code) for one module.

    A COMMENT LINE is one whose first non-whitespace character is the VBA
    apostrophe. A blank line is neither comment nor code. Everything else is code,
    including a continuation line, because VBA charges the reader for it.
    """
    raw = module.raw.splitlines()
    blank = sum(1 for line in raw if not line.strip())
    comment = sum(1 for line in raw if line.strip().startswith("'"))
    return len(raw), blank, comment, len(raw) - blank - comment


def test_05_no_module_is_a_dumping_ground() -> None:
    """The split is by responsibility; one giant module would defeat the point.

    WHY THIS TEST HAS TWO LIMITS NOW. The original Phase-4 assertion was a single
    `raw lines < 900` cap, and it was a PROXY: the thing worth detecting is a
    collapsed responsibility split, and in Phase-4 territory raw size tracked that
    faithfully. It was not a defect - it was a proxy that needed a
    responsibility-aware extension once a module arrived whose responsibility is
    coherent but whose contract requires it to explain itself at length.

    The Phase-4 modules keep the original rule EXACTLY, unrelaxed. The three
    Phase-5 kernel modules are measured on what the proxy was actually reaching
    for - the volume of CODE - and are given a raw ceiling as well, so
    documentation is not charged as sprawl while sprawl is still caught. Neither
    limit alone would do: a code-only limit would let a module grow without bound
    in prose, and a raw-only limit is what penalised documentation in the first
    place.

    The responsibility boundaries themselves are asserted directly in
    tests/test_phase5_vba_source.py; this test is the size half of the pair.
    """
    structure = _specs()[3]
    assert len(structure.vba_modules) >= 6, "the responsibility split collapsed"
    by_name = {m.name: m for m in _handwritten_modules()}
    assert set(by_name) == (
        set(PHASE4_VBA_MODULES) | set(PHASE5_VBA_MODULES) | set(PHASE6_VBA_MODULES)
    ), (
        "the hand-written module inventory changed; the size limits below are "
        "assigned per module and must be assigned for the new one too"
    )
    for name in PHASE4_VBA_MODULES:
        raw, _, _, _ = _line_metrics(by_name[name])
        assert raw < PHASE4_RAW_LINE_LIMIT, (
            f"{name} is {raw} raw lines; split its responsibilities"
        )
    for name in PHASE5_VBA_MODULES + PHASE6_VBA_MODULES:
        raw, _, _, code = _line_metrics(by_name[name])
        assert code < PHASE5_CODE_LINE_LIMIT, (
            f"{name} is {code} code lines; split its responsibilities"
        )
        assert raw < PHASE5_RAW_LINE_LIMIT, (
            f"{name} is {raw} raw lines; split its responsibilities"
        )


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
    # Three declared groups now: button entry points, Phase-4 harness helpers,
    # and the Phase-5 automation/API endpoints. The rule is unchanged - every
    # externally callable PCCM_ procedure is accounted for by the contract.
    accounted = (set(data["vba"]["entry_points"])
                 | set(data["vba"]["harness_procedures"])
                 | set(data["vba"].get("api_procedures", [])))
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
_TYPE_MEMBER_RE = re.compile(r"^\s*(\w+)\s+As\s+\w", re.IGNORECASE)


def _type_member_names(module) -> set[str]:
    """Field names declared inside `Public Type` / `Private Type` blocks."""
    names: set[str] = set()
    inside = False
    for line in module.code_without_string_removal.splitlines():
        stripped = line.strip()
        if re.match(r"^(Public\s+|Private\s+)?Type\s+\w+", stripped, re.IGNORECASE):
            inside = True
            continue
        if re.match(r"^End\s+Type\b", stripped, re.IGNORECASE):
            inside = False
            continue
        if inside:
            match = _TYPE_MEMBER_RE.match(stripped)
            if match:
                names.add(match.group(1))
    return names


def test_11_every_constant_the_vba_references_is_emitted() -> None:
    """The substitute for a VBA compiler.

    A mistyped constant name would otherwise surface only as a Windows runtime
    error, after the review gate. Every SCREAMING_CASE identifier used in the
    hand-written modules must be emitted by a generated module, declared locally,
    or exported as a Public Const by another hand-written module.
    """
    emitted = _generated_constants()
    # A PUBLIC TYPE MEMBER IS A DECLARATION TOO. The scanner reads `Const`
    # declarations, which was the whole of the vocabulary until a module
    # declared a user-defined type: `SimRngState.S10` is SCREAMING_CASE by the
    # regex's reckoning and is defined in the same file that uses it. Missing
    # that is a gap in the scanner, not a missing constant.
    type_members = {
        name for module in _all_modules() for name in _type_member_names(module)
    }
    # VBA and Excel names that are language or library members, not our constants.
    builtin = {
        "VBA", "MSG", "TRUE", "FALSE", "OK", "PCCM", "ID", "URL", "UI",
    }
    problems: list[str] = []
    handwritten = _handwritten_modules()
    exported = {n for module in handwritten for n in _public_constants(module)}
    # A SCREAMING_CASE name may also be defined as a PUBLIC FUNCTION. MAX_DOUBLE
    # is one: a Const initialiser cannot compute, and that boundary has to be
    # built rather than spelled - see modCalcFactors. The rule this test enforces
    # is "referenced and defined somewhere", and a public function defines it
    # just as a Public Const does.
    # PRIVATE procedures count too: TWO_52 and MAX_SIGNIFICAND are private and
    # built, for the same reason MAX_DOUBLE is.
    exported |= {name for module in handwritten for name in module.procedures
                 if name.upper() == name}
    for module in handwritten:
        local = set(module.constants)
        for name in sorted(module.referenced_upper_identifiers):
            if (name in emitted or name in local or name in exported
                    or name in builtin or name in type_members):
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
def test_15_no_forbidden_construct_appears_where_it_is_forbidden() -> None:
    """D6-11 IS PER MODULE, and as of Phase-6 Step 6 that is not academic.

    A bare rule is forbidden everywhere. A scoped rule names the one module
    allowed to contain the construct in executable code, and every other module
    is still forbidden it. Enforcing from the flattened construct list would
    read the scoped construct as global and reject the module that owns it.
    """
    structure = _specs()[3]
    modules = _all_modules()
    problems = []
    for rule in structure.forbidden_construct_rules:
        for module in modules:
            if not rule.forbidden_in(module.name):
                continue
            if contains_construct([module], rule.construct):
                problems.append(f"{rule.construct} in {module.name}")
    assert not problems, "\n".join(problems)


def test_15a_the_one_scoped_grant_is_real_and_is_the_only_one() -> None:
    """The grant is exercised, and it is not a licence for anything else."""
    structure = _specs()[3]
    scoped = [r for r in structure.forbidden_construct_rules if r.is_scoped]
    assert [(r.construct, tuple(r.allowed_in)) for r in scoped] == [
        ("MRG32k3a", ("modSimRng",))
    ], scoped

    modules = {m.name: m for m in _all_modules()}
    assert contains_construct([modules["modSimRng"]], "MRG32k3a"), (
        "the scoped grant is vacuous: modSimRng does not contain the construct "
        "in executable code"
    )
    others = [m for name, m in modules.items() if name != "modSimRng"]
    assert not contains_construct(others, "MRG32k3a")

    # RunSimulation has no owner yet and must not have been scoped early.
    endpoint = [r for r in structure.forbidden_construct_rules
                if r.construct == "RunSimulation"]
    assert len(endpoint) == 1 and not endpoint[0].is_scoped
    for name in modules:
        assert endpoint[0].forbidden_in(name), name


def test_16_no_input_worksheet_change_automation_exists() -> None:
    """Structural operations are command-driven. There is no hidden side effect."""
    modules = _all_modules()
    for construct in ("Worksheet_Change", "Workbook_SheetChange", "Worksheet_SelectionChange"):
        assert not contains_construct(modules, construct), f"{construct} appears in code"


def test_17_no_calculation_or_simulation_code_leaked_in() -> None:
    """Two lists now, because the original list mixed two different prohibitions.

    SIMULATION is forbidden EVERYWHERE, in Phase-5 modules as much as Phase-4 ones:
    no phase that exists yet may draw a random number or take a percentile.

    CALCULATION was forbidden everywhere only because, in Phase 4, everywhere and
    Phase 4 were the same place. Phase 5 is where the calculation belongs, so that
    half of the sweep is retargeted to the Phase-4 modules and the generated
    constants module - which is exactly the territory it was written to protect.
    """
    everywhere = _all_modules()
    # STILL FORBIDDEN EVERYWHERE. No phase that exists yet draws a pseudo-random
    # number, takes a quantile or exposes the simulation endpoint. MRG32k3a has
    # left this list because it now has an owner - it is enforced per module by
    # test_15 and test_15a instead, which is a stronger statement, not a weaker
    # one: it must be in modSimRng and in nothing else.
    for construct in ("Rnd(", "Randomize", "WorksheetFunction.Percentile",
                      "RunSimulation"):
        assert not contains_construct(everywhere, construct), (
            f"{construct} appears in code; no phase that exists yet simulates"
        )
    outside_owner = [m for m in everywhere if m.name != "modSimRng"]
    assert not contains_construct(outside_owner, "MRG32k3a")
    phase4 = [m for m in everywhere
              if m.name not in PHASE5_VBA_MODULES and m.name not in PHASE6_VBA_MODULES]
    for construct in ("ExpectedValue", "DiscountFactor", "EscalationFactor"):
        assert not contains_construct(phase4, construct), (
            f"{construct} appears in Phase-4 code; the calculation lives in the "
            "Phase-5 kernel modules"
        )


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
    emitted = _generated_modules()["modConstants"].constants
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
        "# K2. Profiling PERCENTAGES survive a real reorder, with a LIVE timeline",
        "# A1. The VBA automation surface is callable",
    ):
        assert marker in code, f"the harness is missing section: {marker}"


def test_37a_the_header_matrix_documents_exactly_the_scenarios_that_run() -> None:
    """The header is the first thing a reviewer reads, and it drifted twice.

    Scenario W existed with no matrix entry. Rather than adding one label and moving
    on, the matrix is now derived from what the script actually reports, in both
    directions -- a documented scenario that does not run is as misleading as a
    scenario that runs undocumented.
    """
    text = _ps(HARNESS_PS1)
    header = text[text.index("Test matrix:") : text.index("Safety, unchanged")]
    documented = set(re.findall(r"^\s{6}([A-Z]{1,3}\d?)\s+\S", header, re.MULTILINE))
    reported = set(re.findall(r"Add-Result '([A-Z]{1,3}\d?)'", _ps_code(HARNESS_PS1)))
    # The timeline scenarios are reported from one loop over the oracle fixture as
    # 'D-J.<n>', so their letters are documented individually but never appear as a
    # literal Add-Result label.
    assert "Add-Result ('D-J.' + $stepIndex)" in _ps_code(HARNESS_PS1)
    reported |= set("DEFGHIJ")
    # Housekeeping identifiers the matrix deliberately does not list.
    reported -= {"XX", "Y", "Z"}
    assert documented == reported, (
        f"documented but never run: {sorted(documented - reported)}; "
        f"run but undocumented: {sorted(reported - documented)}"
    )
    assert {"K2", "W", "PRE", "PRE0", "A1"} <= documented


def test_37b_the_first_vba_call_has_its_own_named_boundary() -> None:
    """Importing a module is not compiling it.

    Scenario A imports eight modules, saves, reopens and verifies they persisted --
    and passed on run 3 while the VBA project did not build at all. Excel compiles
    on the first Application.Run, so that call is where a compile error surfaces.
    Without a named step it surfaces inside scenario B and reads as a permanent-ID
    defect, which is the wrong place to start looking.
    """
    code = _ps(HARNESS_PS1)
    marker = "# A1. The VBA automation surface is callable"
    assert marker in code
    section = code[code.index(marker) : code.index("# B. Permanent Cost Line IDs")]

    # It must be the FIRST Application.Run in the script, not merely an early one.
    executable = _ps_code(HARNESS_PS1)
    first_run = executable.index("$excel.Run(")
    assert executable.index("Add-Result 'A1'") > first_run
    assert executable[:first_run].count("Add-Result 'B'") == 0
    assert "PCCM_AutomationBegin" in executable[first_run : first_run + 60], (
        "the first Application.Run must be the automation-surface probe"
    )

    # Real entry points, exercised for real.
    for entry in ("PCCM_AutomationBegin", "PCCM_AutomationResult", "PCCM_AutomationEnd"):
        assert f"$excel.Run('{entry}'" in section, f"A1 must call {entry}"

    # A compile error must surface, never be stepped over or swallowed.
    assert "On Error Resume Next" not in section
    assert "-ErrorAction SilentlyContinue" not in section
    assert "throw" in section, (
        "a failed automation surface must stop the run, not let the scenarios "
        "compare results from a project that never compiled"
    )
    assert "Add-Result 'A1'" in section and "Format-Err $_" in section, (
        "Excel's own compile message must be reported"
    )


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
    # Validation on the grown row is checked against the calibrated Stage-A
    # baseline, through the shared contract helper, not by an ad-hoc probe.
    assert "-Label 'the grown row'" in code
    assert "-Baseline $costValidationBaseline" in code
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


def test_44a1_percentage_ownership_is_proved_with_a_live_timeline() -> None:
    """B2 sorts for real, but before the first Apply -- so there is nothing to own.

    With no project-year column in the profiling grid, B2 can only show that
    identity travels with row data and that profiling rows stay keyed. The claim
    the permanent-ID design exists for -- a PROFILED VALUE belongs to an
    IDENTIFIER, not to a worksheet row -- needs an active timeline.
    """
    code = _ps(HARNESS_PS1)
    marker = "# K2. Profiling PERCENTAGES survive a real reorder"
    assert marker in code, "the live-timeline reorder scenario does not exist"
    section = code[code.index(marker) :]
    section = section[: section.index("# L. Runtime failure containment")]

    # It must run AFTER the timeline scenarios, or it is B2 again.
    assert code.index("# D - J. Timeline scenarios") < code.index(
        "# K2. Profiling PERCENTAGES survive a real reorder"
    ), "K2 must run after a timeline has been applied"

    for check in (
        "at least two identified Cost Lines exist",
        "the profiling grid has at least one project-year column",
        "every seeded percentage is distinct, so a swap cannot pass",
        "the synchronisation pathway ran successfully",
        "the register row order actually changed",
        "every original permanent ID still exists",
        "every profiling percentage still belongs to its own permanent ID",
        "no percentage followed worksheet row position",
        "the profiling grid order follows the reordered register",
        "the ID counter is unchanged by the reorder",
        "structural revalidation is clean after the reorder and sync",
    ):
        assert check in section, f"the K2 scenario is missing: {check}"

    # A REAL sort of whole rows, not an edit in place.
    assert "Invoke-TableSort -Workbook $wb -SheetName $costReg.sheet" in section, (
        "the reorder must be a real ListObject.Sort of the register"
    )
    # The real synchronisation pathway, not a hand-written grid rewrite.
    assert "$excel.Run('PCCM_ApplyTimeline')" in section
    # The comparison is an ID-to-value map captured as plain data beforehand.
    assert "$before[$row[0]] = " in section and "$after[$row[0]] = " in section
    assert "$after[$driverId] -ne $before[$driverId]" in section, (
        "the assertion must compare per identifier, not compare row order"
    )
    assert "$positionalBefore" in section and "$positionalAfter" in section, (
        "position-following must be ruled out explicitly"
    )


# ===========================================================================
# harness result hygiene
#
# Gate-B run 4 reported 18 failures from a handful of causes: one rolled-back
# timeline poisoned five dependent oracle comparisons, one invalid entered triple
# left behind by a deliberate test failed three later scenarios, and one leftover
# Config fixture failed seven more. A failure must be attributable to the thing it
# is testing.
# ===========================================================================
def test_44n1_a_broken_timeline_chain_skips_its_dependents() -> None:
    """D-J is a sequential state machine, and its oracle says so.

    D-J.5's expected headers are computed from the state D-J.4 was supposed to
    leave. Once a step fails, every later comparison is against a state the
    workbook never reached, and reporting those as independent behavioural defects
    is simply untrue.
    """
    code = _ps(HARNESS_PS1)
    assert "$chainBrokenAt = ''" in _ps_code(HARNESS_PS1)
    assert "if ($chainBrokenAt -ne '') {" in _ps_code(HARNESS_PS1)
    assert "'SKIP'" in code[code.index("$chainBrokenAt") :], "dependents must be SKIPped"
    assert "the timeline chain prerequisite failed at" in code
    # Set on any failed step, including one that threw.
    assert _ps_code(HARNESS_PS1).count("$chainBrokenAt = 'D-J.' + $stepIndex") == 2, (
        "the chain must break on a failed checklist AND on an exception"
    )
    # And the run does NOT stop globally: independent scenarios still execute.
    assert code.index("$chainBrokenAt") < code.index("# K. Profiling synchronisation")
    assert "# W. The representation ceiling" in code


def test_44n2_every_timeline_step_asserts_its_outcome_contract() -> None:
    """A rejected Apply looked identical to an accepted one until a later header
    comparison failed, so the reported defect was the header, not the rejection."""
    code = _ps(HARNESS_PS1)
    assert "prevalidation rejected the entered timeline" in code
    assert "the cancelled change reported OK|cancelled" in code
    assert "the accepted change reported OK, not a rejection" in code
    # The exact VBA outcome is printed in every case.
    assert _ps_code(HARNESS_PS1).count('("outcome \'$outcome\'")') >= 3


def test_44n3_scenarios_that_need_a_valid_timeline_normalise_it_first() -> None:
    """D-J.10 deliberately leaves an INVALID entered triple behind.

    Every later scenario that calls Apply for its own purposes was rejected before
    it tested anything -- K2, N and W each failed at run 4 with 'Base Year 2040 is
    later than Project Start Year 2035'. Normalisation is explicit and per
    scenario; it is NOT applied globally, because the entered/applied difference is
    exactly what D-J tests.
    """
    code = _ps_code(HARNESS_PS1)
    helper = _ps_function_body(HARNESS_PS1, "Sync-EnteredTimelineToApplied")
    for pair in (
        ("nmBaseYear_Applied", "nmBaseYear_Entered"),
        ("nmStartYear_Applied", "nmStartYear_Entered"),
        ("nmDuration_Applied", "nmDuration_Entered"),
    ):
        assert pair[0] in helper and pair[1] in helper, pair
    assert "$readBack -ne $applied" in helper, "the copy must be verified, not assumed"

    # Called by each scenario that needs it, and reported as a check so a failed
    # normalisation is visible rather than silent.
    assert code.count("Sync-EnteredTimelineToApplied -Workbook $wb") >= 5
    assert "the entered timeline was normalised to the applied one" in _ps(HARNESS_PS1)
    for marker in ("# K2. Profiling PERCENTAGES", "# N. An Inflation Profile removed",
                   "# T. Unkeyed structural data", "# W. The representation ceiling"):
        section = _ps(HARNESS_PS1)[_ps(HARNESS_PS1).index(marker) :][:6000]
        assert "Sync-EnteredTimelineToApplied" in section, f"{marker} does not normalise"

    # NOT global: the D-J loop must never normalise, or it would erase its own test.
    timeline = _ps(HARNESS_PS1)
    timeline = timeline[timeline.index("# D - J. Timeline scenarios") : timeline.index("# K. Profiling")]
    assert "Sync-EnteredTimelineToApplied" not in timeline


def test_44n4_scenario_o_establishes_its_own_duration_baseline() -> None:
    """O tests destruction of the FINAL profiling year during a shrink, so it needs
    Applied Duration >= 2. D-J.9 deliberately ends at duration 1 and D-J.10 is
    rejected, so inheriting the previous timeline made O impossible to pass in a
    perfectly correct workbook."""
    code = _ps(HARNESS_PS1)
    section = code[code.index("# O. Non-numeric content") : code.index("# P. An oversized")]
    assert "$candidate.expect.applied.duration -ge 2" in section, (
        "the baseline must come from the ORACLE FIXTURE, not be invented here"
    )
    assert "Set-AppliedTimeline" in section
    for check in (
        "the fixture offers a valid baseline with duration >= 2",
        "the baseline timeline applied successfully",
        "the applied triple is the requested baseline",
        "entered equals applied before the shrink",
        "the baseline workbook is structurally clean",
    ):
        assert check in section, f"O is missing its prerequisite check: {check}"
    assert "PASTED TEXT" in section
    assert section.index("the baseline workbook is structurally clean") < section.index("PASTED TEXT"), (
        "the baseline must be established BEFORE the destructive fixture is seeded"
    )


def test_44n5_the_temporary_config_profile_is_always_cleaned_up() -> None:
    """N failed part way at run 4, left 'HARNESS TEMP PROFILE' in Config with no
    matching inflation row, and that one orphan failed every scenario after it."""
    code = _ps(HARNESS_PS1)
    section = code[code.index("# N. An Inflation Profile removed") : code.index("# O. Non-numeric content")]
    assert "} finally {" in section, "the fixture must be removed on the failure path too"
    cleanup = section[section.index("} finally {") :]
    assert "Clear-ConfigProfile" in cleanup
    assert "PCCM_ApplyTimeline" in cleanup, "the removal must be synchronised, not just written"
    assert "PCCM_StructuralReport" in cleanup, "residue must be detected and noted"
    # The profile name is declared outside the try, so the finally can see it.
    assert section.index("$profileName = 'HARNESS TEMP PROFILE'") < section.index("try {")


def test_44n6_independent_scenarios_check_their_prerequisites() -> None:
    """A scenario contaminated by an earlier one is reported as contaminated."""
    code = _ps_code(HARNESS_PS1)
    guard = _ps_function_body(HARNESS_PS1, "Test-CleanStructure")
    assert "PCCM_StructuralReport" in guard
    assert "'SKIP'" in guard, "contamination is a SKIP, not another behavioural failure"
    assert "return $false" in guard and "return $true" in guard
    for scenario in ("P", "Q", "R", "S", "T", "U", "W"):
        assert f"Test-CleanStructure -ExcelApp $excel -ScenarioId '{scenario}'" in code, (
            f"scenario {scenario} does not check that it starts clean"
        )


def test_44n7_the_reorder_scenario_compares_order_not_membership() -> None:
    """Sorting both sides proved membership only: at run 4 Apply was rejected, the
    grid never synchronised, and the assertion passed anyway."""
    code = _ps(HARNESS_PS1)
    section = code[code.index("# K2. Profiling PERCENTAGES") : code.index("# L. Runtime failure")]
    assert "$profileOrderAfter" in section
    assert "($profileOrderAfter -join ',') -eq ($orderAfter -join ',')" in section, (
        "the exact sequence must be compared, with neither side sorted"
    )
    assert "the profiling grid order follows the reordered register EXACTLY" in section
    assert "$after.Keys | Sort-Object" not in section, "the sorted-set comparison must be gone"


def test_44n8_the_orphan_fixture_is_a_real_blank_row() -> None:
    """Add-then-Delete cannot make a free row: DeleteDriver removes the ListRow, so
    the table returns to the same count and the fixture wrote at row index zero."""
    code = _ps(HARNESS_PS1)
    section = code[code.index("# T. Unkeyed structural data") : code.index("# U. A corrupt ID counter")]
    assert "Add-BlankTableRow" in section, "the fixture must add a blank row directly"
    # Only the CREATION of the fixture is constrained -- calling Add afterwards is
    # the point of the scenario, since Add is what must be refused.
    code_only = _ps_code(HARNESS_PS1)
    creation = code_only[code_only.index("$fixtureRow = Add-BlankTableRow") :]
    creation = creation[: creation.index("ORPHAN DESCRIPTION")]
    assert "PCCM_AddCostLine" not in creation, (
        "the corruption fixture must not allocate a permanent ID"
    )
    assert "PCCM_DeleteCostLineById" not in creation, (
        "Add-then-Delete is exactly what could not produce a free row"
    )
    for check in (
        "a blank register row was added for the corruption fixture",
        "the fixture row exists in the table body",
        "the fixture row has no permanent ID",
        "every cell of the fixture row is blank",
        "Add is refused while a driver orphan exists",
        "the orphan row is untouched",
    ):
        assert check in section, f"the T1 fixture is missing: {check}"
    assert "Remove-TableRow" in section, "the fixture row must be removed afterwards"
    # And the helper owns its COM leaf-before-parent, like every other reader.
    adder = _ps_function_body(HARNESS_PS1, "Add-BlankTableRow")
    assert "Release-Transient $added" in adder and "Release-Transient $rows" in adder


def test_44n9_validation_is_compared_against_a_calibrated_baseline() -> None:
    """The old check read $cell.Validation and treated "no exception" as "there is
    a user restriction". Excel returns a Validation object either way, and a Type
    of xlValidateInputOnly restricts nothing."""
    code = _ps(HARNESS_PS1)
    assert "Test-TableCellValidation" not in _ps_code(HARNESS_PS1), (
        "the assumption-based probe must be gone, not merely unused"
    )
    fingerprint = _ps_function_body(HARNESS_PS1, "Get-ValidationFingerprint")
    assert "$validation.Type" in fingerprint
    assert "$validation.Formula1" in fingerprint and "$validation.Formula2" in fingerprint

    section = code[code.index("# A2. Data Validation baseline") : code.index("# B. Permanent Cost Line IDs")]
    assert "-RowIndex 1" in section, "the baseline comes from an untouched Stage-A row"
    assert section.index("Get-RowValidationFingerprints") < code.index("PCCM_AddCostLine"), (
        "the baseline must be captured before any driver is added"
    )
    assert "the fingerprint distinguishes a validated column from the ID column" in section, (
        "a baseline of 'nothing anywhere' would make every later comparison vacuous"
    )

    # Applied to the grown row and to both rollback rows.
    assert _ps_code(HARNESS_PS1).count("-Baseline $costValidationBaseline") == 3
    assert "ID cell has no constraining user validation, matching the model-controlled baseline" in code, (
        "the ID assertion must be phrased as the contract states it"
    )
    assert "carries NO user Data Validation" not in code, (
        "the old phrasing encoded the wrong assumption"
    )
    # Coverage is not weakened: the validated user columns are still named.
    assert "every validated user column keeps its baseline Data Validation" in code


def test_44n10_the_excel_identity_variable_is_never_clobbered() -> None:
    """$id held the Excel process identity AND was reused as a foreach iterator.

    PowerShell loop variables are not block-scoped, so by shutdown $id was a driver
    identifier string, and Wait-ExcelExit / $id.ProcessId failed with
    PropertyNotFoundException at run 4.
    """
    for path in (BUILD_PS1, HARNESS_PS1):
        code = _ps_code(path)
        identities = set(re.findall(r"\$(\w+)\s*=\s*Get-ExcelIdentity", code))
        assert identities, f"{path.name}: no Excel identity is captured"
        for name in identities:
            assert "identity" in name.lower(), (
                f"{path.name}: ${name} holds the Excel identity but is not named for it"
            )
            assert not re.search(rf"foreach\s*\(\s*\${name}\s+in\b", code), (
                f"{path.name}: ${name} is reused as a loop iterator"
            )
            others = [
                line.strip()
                for line in code.splitlines()
                if re.match(rf"\s*\${name}\s*=", line)
                and "Get-ExcelIdentity" not in line
                and "= $null" not in line
            ]
            assert not others, f"{path.name}: ${name} is overwritten by {others}"
        # Every consumer uses the identity variable, not a bare $id.
        for consumer in ("Wait-ExcelExit -Identity", "Invoke-EmergencyExcelCleanup -Identity"):
            for line in code.splitlines():
                if consumer in line:
                    assert re.search(r"-Identity\s+\$\w*[Ii]dentity\b", line), line.strip()

    # And the loops that caused it now carry semantic names.
    harness = _ps_code(HARNESS_PS1)
    assert re.search(r"foreach\s*\(\s*\$driverId\s+in\b", harness)
    assert re.search(r"foreach\s*\(\s*\$riskId\s+in\b", harness)
    assert not re.search(r"foreach\s*\(\s*\$id\s+in\b", harness)


def test_44o1_no_fixture_writes_into_an_unkeyed_profiling_row() -> None:
    """A test fixture must not itself break the invariant under test.

    O neutralised the tail project year by looping every physical body row of both
    profiling grids and writing numeric zero -- including the UNKEYED reserved
    rows. That is exactly the orphan the model refuses: blank key plus data
    elsewhere in the row. At run 5 PreMutationCheck stopped Apply before the
    destructive assessment ran, so O reported a missing prompt while the real cause
    was its own fixture, and the residue skipped P, Q, R, S, T, U and W.

    Numeric zero is non-destructive only in a KEYED cell. Every profiling or
    register write must therefore sit under a permanent-ID guard -- except the
    deliberate orphan fixtures in T, which are the whole point of that scenario.
    """
    lines = _ps_structural_lines(HARNESS_PS1)
    raw = _ps_code(HARNESS_PS1).splitlines()
    commented = _ps(HARNESS_PS1).splitlines()

    # Which scenario each line belongs to. Read from the COMMENTED source, since
    # _ps_code blanks comment lines; both keep the same line numbering.
    sections: list[tuple[int, str]] = []
    for number, line in enumerate(commented, 1):
        match = re.match(r"\s*# ([A-Z]{1,3}\d?)\. ", line)
        if match:
            sections.append((number, match.group(1)))
    assert len(sections) > 15, f"only {len(sections)} scenario markers found"

    # Helper bodies are generic utilities; the permanent-ID guard belongs at the
    # CALL SITE, so a write inside a function definition is not judged here.
    in_function: set[int] = set()
    index = 0
    while index < len(lines):
        if re.match(r"\s*function\s+[\w-]+", lines[index]):
            depth, opened, end = 0, False, index
            for scan in range(index, len(lines)):
                for char in lines[scan]:
                    if char == "{":
                        depth += 1
                        opened = True
                    elif char == "}":
                        depth -= 1
                end = scan
                if opened and depth == 0:
                    break
            in_function.update(range(index + 1, end + 2))
            index = end
        index += 1

    def scenario_at(number: int) -> str:
        current = "setup"
        for start, name in sections:
            if start <= number:
                current = name
            else:
                break
        return current

    # A write is guarded when a permanent-ID test governs it: either an explicit
    # key test in the enclosing lines, or a row index resolved by matching a key.
    guard = re.compile(
        r"\$row\[0\]\s*-ne\s*''"          # the row is identified
        r"|\$row\[0\]\s*-eq\s+\$"          # ... or matched against a specific id
        r"|ContainsKey\(\$row\[0\]\)"
        r"|\$seedRow\s*-gt\s*0|\$blankRow\s*-gt\s*0|\$gridRow\s*-gt\s*0"
        r"|\$profileRow\s*-gt\s*0|\$targetRow\s*-gt\s*0"
        r"|\$key\s*-ne\s*''"
    )
    unguarded = []
    for number, line in enumerate(lines, 1):
        if "Set-TableCell -Workbook" not in line:
            continue
        if number in in_function:
            continue
        scenario = scenario_at(number)
        if scenario == "T":
            continue  # the deliberate orphan fixtures
        window = "\n".join(raw[max(0, number - 12) : number + 1])
        if not guard.search(window):
            unguarded.append(f"{scenario} @ {number}: {raw[number - 1].strip()[:70]}")
    assert not unguarded, (
        "structural-table write with no permanent-ID guard:\n  " + "\n  ".join(unguarded)
    )

    # And the exact run-5 shape is rejected by name: a blind loop over every body
    # row of a profiling grid, writing a value.
    blind = re.compile(
        r"for \(\$r = 1; \$r -le \(Get-TableRowCount[^)]*(?:costGrid|riskGrid)"
    )
    assert not blind.search(_ps_code(HARNESS_PS1)), (
        "a blind row loop over a profiling grid is how run 5 wrote into reserved rows"
    )


def test_44o2_the_o_fixture_is_keyed_only_and_uses_each_grids_own_width() -> None:
    """Per-grid fixed column counts, not the coincidence that both are two."""
    code = _ps(HARNESS_PS1)
    section = code[code.index("# O. Non-numeric content") : code.index("# P. An oversized")]
    assert "$fixedRisk  = $riskGrid.fixed_columns.Count" in _ps_code(HARNESS_PS1), (
        "the Risk grid must be indexed by its own contract-derived fixed count"
    )
    assert "$oTailCost = $fixedCost + $oDuration" in section
    assert "$oTailRisk = $fixedRisk + $oDuration" in section
    assert "$fixedCost + [int]$appliedDuration" not in section, (
        "the old shared-width indexing must be gone"
    )
    # Both zeroing loops are keyed-only.
    zeroing = section[section.index("neutralise the tail year") :]
    zeroing = zeroing[: zeroing.index("the one genuinely destructive cell")]
    assert zeroing.count("if ($row[0] -ne '') {") == 2, (
        "both profiling grids must be zeroed under a permanent-ID guard"
    )
    assert "-Value 0" in zeroing


def test_44o3_the_o_fixture_is_proved_structurally_clean_before_apply() -> None:
    """PASTED TEXT in a KEYED cell is business-invalid but not a structural orphan.

    O needs the workbook coherent so Apply reaches the destructive assessment
    rather than stopping at PreMutationCheck -- which is what the missing prompt
    was really telling us at run 5.
    """
    code = _ps(HARNESS_PS1)
    section = code[code.index("# O. Non-numeric content") : code.index("# P. An oversized")]
    assert "the O fixture is structurally clean before the destructive assessment" in section
    apply_at = section.index("$excel.Run('PCCM_ApplyTimeline')")
    check_at = section.index("the O fixture is structurally clean before the destructive assessment")
    paste_at = section.index("-Value 'PASTED TEXT'")
    duration_at = section.index("Set-NamedValue -Workbook $wb -DefinedName 'nmDuration_Entered'")
    assert paste_at < check_at and duration_at < check_at, (
        "the fixture must be complete before it is checked"
    )
    assert check_at < apply_at, "the check must precede Apply"


def test_44o4_the_o_outcome_is_asserted_not_inferred_from_the_prompt() -> None:
    """Reading only the prompt hid why it was empty: a refusal reports itself in
    the outcome, and reporting 'no prompt' for a refused operation names the wrong
    defect."""
    code = _ps(HARNESS_PS1)
    section = code[code.index("# O. Non-numeric content") : code.index("# P. An oversized")]
    assert "$outcome = [string]$excel.Run('PCCM_AutomationResult')" in section
    assert "the operation reached the confirmation and was cancelled" in section
    assert "($outcome -eq 'OK|cancelled')" in section, (
        "the cancellation contract must be exact"
    )
    outcome_at = section.index("the operation reached the confirmation and was cancelled")
    prompt_at = section.index("text in a removed profiling cell triggers a destructive warning")
    assert outcome_at < prompt_at, "the outcome must be reported before the prompt claims"
    assert "PERMANENTLY DELETED" in section
    assert "the affected permanent ID is named" in section


def test_44o5_the_o_fixture_is_failure_safe_and_restores_exact_values() -> None:
    """Run 5 proved a scenario's disposable data outliving a FAILED run
    contaminates every later independent scenario."""
    code = _ps(HARNESS_PS1)
    section = code[code.index("# O. Non-numeric content") : code.index("# P. An oversized")]
    assert "} finally {" in section, "cleanup must run on the failure path too"
    cleanup = section[section.index("--- CLEANUP, WHATEVER HAPPENED") :]

    # Captured by permanent ID, never by physical row position.
    assert "$oTouchedCost[[string]$row[0]] = $value" in section
    assert "$oTouchedRisk[[string]$row[0]] = $value" in section
    # A blank comes back blank, never as numeric zero.
    assert "if ($original -eq '') {" in cleanup and "-Value $null" in cleanup
    assert "-Value ([double]$original)" in cleanup
    # Entered timeline restored, and no synchronisation called just to erase residue.
    assert "Sync-EnteredTimelineToApplied" in cleanup
    assert "PCCM_ApplyTimeline" not in cleanup, (
        "a structural synchronisation must not be used to clean up fixture residue"
    )
    # Cleanup failure keeps the gate failed, and is noted.
    assert "the O fixture cleanup completed without error" in cleanup
    assert "Add-Note ('O cleanup problems: '" in cleanup
    assert "structural revalidation is clean after O fixture cleanup" in cleanup
    # The result is reported AFTER cleanup, so cleanup failures reach the verdict.
    assert cleanup.index("the O fixture cleanup completed without error") < cleanup.index(
        "Add-Result 'O'"
    )


def test_44o6_o_proves_the_unkeyed_rows_were_never_touched() -> None:
    """The strong claim is that O never writes to them -- not that it tidies up
    afterwards."""
    code = _ps(HARNESS_PS1)
    section = code[code.index("# O. Non-numeric content") : code.index("# P. An oversized")]
    assert "the reserved unkeyed profiling rows start blank" in section
    assert "the fixture wrote to no unkeyed profiling row" in section
    assert "every unkeyed profiling tail cell is unchanged after cleanup" in section
    # Captured before the fixture is built, and checked twice.
    capture_at = section.index("$oUnkeyedCost[$rowIdx] = $value")
    build_at = section.index("neutralise the tail year")
    first_check = section.index("the fixture wrote to no unkeyed profiling row")
    last_check = section.index("every unkeyed profiling tail cell is unchanged after cleanup")
    assert capture_at < build_at < first_check < last_check
    # And the later independent gates are still there.
    assert "Test-CleanStructure -ExcelApp $excel -ScenarioId 'P'" in _ps_code(HARNESS_PS1)


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
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1, SCENARIOS_PS1):
        defined |= set(re.findall(r"^\s*function\s+([\w-]+)", _ps(path), re.MULTILINE))

    # Cmdlets and functions provided by PowerShell itself, used deliberately.
    builtin = {
        "Get-Content", "Set-Content", "Get-Process", "Stop-Process", "Get-Date",
        "Start-Sleep", "Write-Host", "Test-Path", "Remove-Item", "New-Item",
        "Copy-Item", "Join-Path", "Split-Path", "New-Object", "Add-Type",
        "Select-Object", "Where-Object", "ForEach-Object", "ConvertFrom-Json",
        "Set-StrictMode", "Get-CimInstance", "Out-Null", "Write-Verbose",
        "Sort-Object", "Measure-Object", "Select-String", "Write-Output",
    }
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1, SCENARIOS_PS1):
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


_VBA_DECLARATION = re.compile(
    r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+\w+)\s+(\w+)",
    re.IGNORECASE,
)


def _vba_declarations() -> list[tuple[str, str, str]]:
    """(module, procedure, joined declaration) for every Sub/Function/Property.

    Continuations are joined first: a parameter list wrapped across lines is one
    declaration, and reading only the first physical line would miss every
    parameter after the underscore.
    """
    found = []
    for module in _all_modules():
        lines = module.code_without_string_removal.splitlines()
        index = 0
        while index < len(lines):
            match = _VBA_DECLARATION.match(lines[index])
            if match:
                declaration = lines[index].rstrip()
                last = index
                while declaration.endswith("_"):
                    last += 1
                    declaration = declaration[:-1].rstrip() + " " + lines[last].strip()
                found.append((module.name, match.group(1), " ".join(declaration.split())))
                index = last
            index += 1
    return found


def _vba_parameters(declaration: str) -> list[str]:
    """The parameter list, split on commas that are not inside parentheses."""
    start = declaration.find("(")
    if start < 0:
        return []
    depth, end = 0, len(declaration)
    for position in range(start, len(declaration)):
        if declaration[position] == "(":
            depth += 1
        elif declaration[position] == ")":
            depth -= 1
            if depth == 0:
                end = position
                break
    inner = declaration[start + 1 : end]
    parts, depth, current = [], 0, ""
    for char in inner:
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        if char == "," and depth == 0:
            parts.append(current)
            current = ""
        else:
            current += char
    if current.strip():
        parts.append(current)
    return [p.strip() for p in parts if p.strip()]


_VBA_PROCEDURE_START = re.compile(
    r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+",
    re.IGNORECASE,
)
_VBA_PROCEDURE_END = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)

# Module-level declaration forms. `Dim` at module level is legal VBA and belongs in
# the declaration section too; inside a procedure it is the ordinary local form and
# must NOT be flagged, which is why this is driven by procedure boundaries rather
# than by the keyword alone.
_VBA_MODULE_DECLARATION = re.compile(
    r"^(?:"
    r"(?:Public|Private|Global|Dim)\s+"
    r"(?!Sub\b|Function\b|Property\b|Declare\b|Type\b|Enum\b|Const\b)\w+"
    r"|(?:Public\s+|Private\s+)?Const\s+\w+"
    r"|(?:Public\s+|Private\s+)?(?:Type|Enum)\s+\w+"
    r")",
    re.IGNORECASE,
)


def _declaration_section_offenders(module) -> list[tuple[int, str]]:
    """Module-level declarations that appear AFTER the first executable procedure."""
    offenders: list[tuple[int, str]] = []
    first_procedure_seen = False
    inside_procedure = False
    for number, statement in logical_statements(module.code_without_string_removal):
        if _VBA_PROCEDURE_START.match(statement):
            first_procedure_seen = True
            inside_procedure = True
            continue
        if _VBA_PROCEDURE_END.match(statement):
            inside_procedure = False
            continue
        if inside_procedure:
            continue  # Dim / Const / Static inside a procedure are ordinary locals
        if first_procedure_seen and _VBA_MODULE_DECLARATION.match(statement):
            offenders.append((number, statement))
    return offenders


def test_45f_every_module_level_declaration_is_in_the_declaration_section() -> None:
    """VBA has no "module-level statement anywhere in the file".

    Everything before the first executable procedure is the declaration section;
    everything after it is procedure bodies. A Public variable, Const, Type or Enum
    written after that point is not merely untidy -- under Option Explicit the
    compiler reaches the procedure that uses it with the name undefined and stops.

    That is what ended Gate-B run 3, on the harness's FIRST VBA call: the five
    gAutomation* globals sat after ConfirmDestructiveChange, and VBA raised
    "Compile error: Variable not defined" on gAutomationActive. This sweep is the
    standing substitute for the compiler that cannot run on Linux.
    """
    problems = []
    for module in _all_modules():
        if "option explicit" not in module.code.lower():
            continue
        for number, statement in _declaration_section_offenders(module):
            problems.append(f"{module.name}:{number}: {statement[:70]}")
    assert not problems, (
        "module-level declaration after the first executable procedure "
        "(Compile error: Variable not defined):\n  " + "\n  ".join(problems)
    )


def test_45g_procedure_local_declarations_are_not_flagged() -> None:
    """The sweep must not be a grep for `Dim` -- locals are the common case.

    A sweep that flagged procedure-local Dim, Const or Static would be unusable and
    would be deleted, which is a worse outcome than not having it. This proves the
    procedure-boundary tracking works, on real source and on a constructed module
    that is deliberately legal.
    """
    workbook = next(m for m in _all_modules() if m.name == "modWorkbook")
    locals_found = [
        statement
        for _, statement in logical_statements(workbook.code_without_string_removal)
        if re.match(r"(Dim|Const|Static)\s+\w+", statement, re.IGNORECASE)
    ]
    assert len(locals_found) >= 15, (
        f"expected many procedure-local declarations, found {len(locals_found)}"
    )
    assert not _declaration_section_offenders(workbook)

    class _Fake:
        name = "modFake"
        code = "option explicit"
        code_without_string_removal = (
            "Option Explicit\n"
            "Public gFlag As Boolean\n"
            "Private Const K As Long = 1\n"
            "Public Type T\n"
            "    A As Long\n"
            "End Type\n"
            "Public Sub S()\n"
            "    Dim local As Long\n"
            "    Const inner As Long = 2\n"
            "    Static kept As Long\n"
            "    local = inner + kept\n"
            "End Sub\n"
            "Public Function F() As Long\n"
            "    Dim other As String\n"
            "    F = 1\n"
            "End Function\n"
        )

    assert _declaration_section_offenders(_Fake()) == [], (
        "legal procedure-local declarations were flagged"
    )

    class _Broken(_Fake):
        name = "modBroken"
        code_without_string_removal = _Fake.code_without_string_removal + (
            "Public gLate As Boolean\n"
            "Public Const LATE_K As Long = 3\n"
            "Public Type TLate\n"
            "    B As Long\n"
            "End Type\n"
        )

    caught = [statement for _, statement in _declaration_section_offenders(_Broken())]
    assert caught == [
        "Public gLate As Boolean",
        "Public Const LATE_K As Long = 3",
        "Public Type TLate",
    ], caught


def test_45h_the_sweep_joins_line_continuations() -> None:
    """A wrapped declaration is still one declaration.

    Per-physical-line analysis reads the tail of a wrapped statement as a statement
    of its own, which both misses real offenders and invents false ones.
    """
    statements = logical_statements(
        "Public gWrapped _\n"
        "    As Boolean\n"
        "Public Sub S()\n"
        "End Sub\n"
        "Public gAfter _\n"
        "    As Long\n"
    )
    assert (1, "Public gWrapped As Boolean") in statements
    assert (5, "Public gAfter As Long") in statements

    class _Wrapped:
        name = "modWrapped"
        code = "option explicit"
        code_without_string_removal = (
            "Option Explicit\n"
            "Public Sub S()\n"
            "End Sub\n"
            "Public gAfter _\n"
            "    As Long\n"
        )

    assert [s for _, s in _declaration_section_offenders(_Wrapped())] == [
        "Public gAfter As Long"
    ]


def test_45i_the_automation_globals_are_in_the_declaration_section() -> None:
    """The five that actually failed, pinned by name and by position."""
    module = next(m for m in _handwritten_modules() if m.name == "modAppState")
    statements = logical_statements(module.code_without_string_removal)
    first_procedure = next(
        number for number, statement in statements if _VBA_PROCEDURE_START.match(statement)
    )
    expected = [
        "gAutomationActive",
        "gAutomationConfirmReply",
        "gAutomationLastPrompt",
        "gAutomationFailAfterStage",
        "gAutomationLastResult",
    ]
    for name in expected:
        declarations = [
            number
            for number, statement in statements
            if re.match(rf"Public\s+{name}\s+As\s+\w+$", statement)
        ]
        assert len(declarations) == 1, (
            f"{name} is declared {len(declarations)} times; it must be declared once"
        )
        assert declarations[0] < first_procedure, (
            f"{name} is declared at line {declarations[0]}, after the first "
            f"executable procedure at line {first_procedure}"
        )
    # Public, not Private: modTimeline and modDrivers reference them by module name.
    for other in ("modTimeline", "modDrivers"):
        code = next(m for m in _handwritten_modules() if m.name == other).code
        assert "modAppState.gAutomation" in code, (
            f"{other} references the automation state, so it must stay Public"
        )


def test_45c_no_typed_optional_parameter_omits_its_default() -> None:
    """`Optional ByRef Unrepresentable As Long` does not compile.

    VBA allows an Optional parameter with no default ONLY when its type is Variant,
    because Variant can hold Missing. A typed Optional must declare a default value.
    This one shipped as a Gate-B compile blocker that no Linux test could see, so the
    sweep is permanent and covers every declaration, not just that procedure.

    Legitimate forms stay legal: a Variant Optional (explicit or implicit) and a typed
    Optional that declares a default.
    """
    problems = []
    for module, procedure, declaration in _vba_declarations():
        for parameter in _vba_parameters(declaration):
            if not re.match(r"Optional\b", parameter, re.IGNORECASE):
                continue
            if "=" in parameter:
                continue  # a default is declared
            type_clause = re.search(r"\bAs\s+(\w+)", parameter, re.IGNORECASE)
            if type_clause is None:
                continue  # implicit Variant
            if type_clause.group(1).lower() == "variant":
                continue
            problems.append(f"{module}.{procedure}: {parameter}")
    assert not problems, (
        "typed Optional parameter with no default (a VBA compile error):\n  "
        + "\n  ".join(problems)
    )


def test_45d_highest_issued_takes_a_required_out_parameter() -> None:
    """Every caller supplies it, so optional semantics bought nothing and cost a build."""
    declaration = next(
        text
        for module, procedure, text in _vba_declarations()
        if module == "modDrivers" and procedure == "HighestIssued"
    )
    parameters = _vba_parameters(declaration)
    assert parameters == ["ByVal Kind As String", "ByRef Unrepresentable As Long"], parameters
    # And the corruption count is genuinely consumed, not merely accepted.
    check = next(m for m in _handwritten_modules() if m.name == "modStructuralCheck").code
    assert "modDrivers.HighestIssued(kind, unrepresentable)" in check


def test_45e_the_declaration_sweep_reads_the_whole_parameter_list() -> None:
    """A sweep that stopped at the first physical line would pass on anything.

    Almost every Phase-4 declaration wraps, HighestIssued included, so the joining
    step is the part that can silently make this whole test inert.
    """
    declarations = _vba_declarations()
    assert len(declarations) > 100, f"only {len(declarations)} declarations found"
    wrapped = [
        (module, procedure)
        for module, procedure, text in declarations
        if len(_vba_parameters(text)) > 1
    ]
    assert len(wrapped) > 20, "the parameter splitter is not seeing multi-parameter lists"
    assert ("modDrivers", "HighestIssued") in wrapped
    dangling = [
        f"{module}.{procedure}"
        for module, procedure, text in declarations
        if text.rstrip().endswith("_")
        or any(p.rstrip().endswith("_") for p in _vba_parameters(text))
    ]
    assert not dangling, f"a continuation was left unjoined: {dangling}"


# ===========================================================================
# collision-safe header renaming
#
# Excel requires ListObject column names to be unique and does NOT refuse a
# collision -- it silently appends a digit. A single sequential pass therefore
# corrupts any OVERLAPPING rename, which is the common case here: shifting the
# start year asks 2028..2032 -> 2030..2034, and the target machine came back with
# 20272, 20282, 20292, 20302 at Gate-B run 4.
# ===========================================================================
def _rename_sequentially(current: list[str], final: list[str]) -> list[str]:
    """A MODEL of Excel's behaviour, not an implementation of the fix.

    Renaming to a name another column still holds does not fail; Excel appends a
    digit. This exists so the two-pass algorithm can be shown to be necessary
    rather than asserted to be.
    """
    names = list(current)
    for index, wanted in enumerate(final):
        taken = {n for i, n in enumerate(names) if i != index}
        candidate, suffix = wanted, 1
        while candidate in taken:
            suffix += 1
            candidate = f"{wanted}{suffix}"
        names[index] = candidate
    return names


def _rename_two_pass(current: list[str], final: list[str]) -> list[str]:
    """The algorithm the VBA implements: vacate every name, then place them."""
    names = list(current)
    temps = []
    for index in range(len(final)):
        suffix, candidate = 0, f"PCCM_TMP_HDR_{index + 1}"
        while candidate in names or candidate in final or candidate in temps:
            suffix += 1
            candidate = f"PCCM_TMP_HDR_{index + 1}_{suffix}"
        temps.append(candidate)
    for index, temp in enumerate(temps):
        names = _rename_sequentially(names, [temp if i == index else n for i, n in enumerate(names)])
    for index, wanted in enumerate(final):
        names = _rename_sequentially(names, [wanted if i == index else n for i, n in enumerate(names)])
    return names


def test_45j_a_sequential_rename_provably_corrupts_an_overlapping_block() -> None:
    """The defect, demonstrated rather than described.

    If this ever stops failing, the model of Excel's behaviour has drifted and the
    two-pass test below would be proving nothing.
    """
    before = ["2028", "2029", "2030", "2031", "2032"]
    after = ["2030", "2031", "2032", "2033", "2034"]
    corrupted = _rename_sequentially(before, after)
    assert corrupted != after, "a sequential rename of an overlapping block must corrupt"
    assert any(name not in after for name in corrupted), corrupted

    # Whether a sequential pass collides is DIRECTION-DEPENDENT, and that is the
    # trap: shifting the same block back DOWN, left to right, happens to vacate
    # each target before it is needed, so it survives by luck.
    assert _rename_sequentially(after, before) == before

    # Luck runs out on any rollback that is not a monotonic downward shift. A
    # reversal -- which RestoreTable can face, since it restores a whole header row
    # to whatever the snapshot held -- collides immediately.
    reversal = _rename_sequentially(["2028", "2029", "2030"], ["2030", "2029", "2028"])
    assert reversal != ["2030", "2029", "2028"], reversal


def test_45k_the_two_pass_rename_is_correct_in_both_directions() -> None:
    """Forward shift, rollback shift, full reversal and a no-op all land exactly."""
    cases = [
        (["2028", "2029", "2030", "2031", "2032"], ["2030", "2031", "2032", "2033", "2034"]),
        (["2030", "2031", "2032", "2033", "2034"], ["2028", "2029", "2030", "2031", "2032"]),
        (["2028", "2029", "2030"], ["2030", "2029", "2028"]),
        (["2028", "2029", "2030"], ["2028", "2029", "2030"]),
        (["2035"], ["2036"]),
    ]
    for before, after in cases:
        assert _rename_two_pass(before, after) == after, (before, after)


def test_45l_one_collision_safe_primitive_exists_and_is_the_only_renamer() -> None:
    """Three paths mutated headers; one primitive now does it for all of them."""
    workbook = next(m for m in _handwritten_modules() if m.name == "modWorkbook")
    body = workbook.code[workbook.code.index("Public Sub SetHeaderBlock") :]
    body = body[: body.index("End Sub")]

    # The required algorithm, step by step.
    assert "ReDim temps(1 To total)" in body, "temporary names must be materialised"
    assert "HEADER_TEMP_PREFIX" in body, "deterministic temporary names, not random"
    assert "Rnd" not in body and "Random" not in body
    assert "HeaderNameInUse(Target, candidate)" in body, "checked against current names"
    assert "NameInArray(FinalNames, candidate, total)" in body, "checked against desired names"
    assert "NameInArray(temps, candidate, i - 1)" in body, "checked against each other"
    first_pass = body.index("Target.ListColumns(FirstColumn + i - 1).Name = temps(i)")
    second_pass = body.index("Target.ListColumns(FirstColumn + i - 1).Name = FinalNames(i)")
    verify = body.index("vbBinaryCompare")
    assert first_pass < second_pass < verify, "vacate, then place, then verify"
    assert "StrComp" in body[verify - 200 : verify + 200]
    assert "Err.Raise" in body[verify:], "a differing final name must raise, not be accepted"

    # Duplicate desired names can never succeed, so they are refused up front.
    assert "would both be" in workbook.code_without_string_removal

    # AND NO OTHER RENAMER SURVIVES. Both a ListColumn.Name assignment and a header
    # CELL write rename a column, and both collide the same way.
    problems = []
    for module in _handwritten_modules():
        for number, line in enumerate(module.code.splitlines(), 1):
            if re.search(r"ListColumns\([^)]*\)\.Name\s*=", line) and module.name != "modWorkbook":
                problems.append(f"{module.name}:{number}: {line.strip()[:60]}")
            if re.search(r"HeaderRowRange\.Cells\([^)]*\)\.Value\s*=", line):
                problems.append(f"{module.name}:{number}: {line.strip()[:60]}")
    assert not problems, (
        "an independent header rename survives outside the primitive:\n  "
        + "\n  ".join(problems)
    )
    # Inside modWorkbook the only two assignments are the primitive's own passes.
    renames = [
        line.strip()
        for line in workbook.code.splitlines()
        if re.search(r"ListColumns\([^)]*\)\.Name\s*=", line)
    ]
    assert len(renames) == 2, renames


def test_45m_all_three_header_mutating_paths_use_the_primitive() -> None:
    """Profiling, inflation and ROLLBACK. The rollback is not an afterthought:
    restoring 2030..2034 to 2028..2032 overlaps exactly as badly, and a restore
    that corrupts the headers it is putting back is not a restore."""
    for name, marker in (
        ("modProfiling", "modWorkbook.SetHeaderBlock target, fixedCols + 1, wantedHeaders"),
        ("modInflation", "modWorkbook.SetHeaderBlock target, fixedCols + 1, wantedHeaders"),
        ("modWorkbook", "SetHeaderBlock Target, 1, restoredHeaders"),
    ):
        code = next(m for m in _handwritten_modules() if m.name == name).code
        assert marker in code, f"{name} does not route its rename through the primitive"

    restore = next(m for m in _handwritten_modules() if m.name == "modWorkbook").code
    restore = restore[restore.index("Public Sub RestoreTable") :]
    restore = restore[: restore.index("End Sub")]
    assert "SetHeaderBlock Target, 1, restoredHeaders" in restore
    assert "Target.HeaderRowRange.Cells(1, c).Value = Snapshot.Headers(c)" not in restore, (
        "the per-cell header write is the same collision-unsafe rename"
    )


def test_45n_the_harness_reports_header_residue_by_position() -> None:
    """Exact equality per position, with no assumption about the bad suffix."""
    code = _ps(HARNESS_PS1)
    assert "every header is EXACTLY the requested value, with no rename residue" in code
    assert "-cne [string]$pair.Expected[$h]" in _ps_code(HARNESS_PS1), (
        "the comparison must be exact and case-sensitive, per position"
    )
    assert "'2'" not in _ps_code(HARNESS_PS1).replace("'20", "'XX"), (
        "no particular collision suffix may be hardcoded as the only bad outcome"
    )


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


# Helpers whose result is a COLLECTION: zero items is a valid outcome, and a
# PowerShell function returning an empty collection emits ZERO pipeline objects. The
# assignment therefore lands $null however the helper wrote its return, and under
# Set-StrictMode reading .Count on $null raises PropertyNotFoundException. That is
# not theoretical: it ended the first real Gate-B run, on the SUCCESS path, because
# zero transient release failures is exactly what produces it.
COLLECTION_HELPERS = (
    "Get-TransientFailures",
    "Get-PreExistingExcelPids",
    "Get-IdColumnValues",
    "Get-TableBody",
    "Get-TableColumnNames",
)

# Of those, the one that returns a JAGGED collection. Caller-side @(...) cannot
# repair this one alone -- the helper must also suppress enumeration on output.
JAGGED_COLLECTION_HELPERS = ("Get-TableBody",)


def _ps_function_body(path: Path, name: str) -> str:
    """The body of one PowerShell function, delimited by brace depth."""
    lines = _ps_structural_lines(path)
    raw = _ps_code(path).splitlines()
    start = next(
        (i for i, line in enumerate(lines) if re.match(rf"\s*function\s+{re.escape(name)}\b", line)),
        None,
    )
    assert start is not None, f"{path.name}: no function {name}"
    depth, opened, end = 0, False, len(lines)
    for index in range(start, len(lines)):
        for char in lines[index]:
            if char == "{":
                depth += 1
                opened = True
            elif char == "}":
                depth -= 1
        if opened and depth == 0:
            end = index + 1
            break
    return "\n".join(raw[start:end])


def test_46l_the_transient_failure_report_materialises_its_collection() -> None:
    """The exact defect that ended the first Gate-B run, pinned in both scripts."""
    for path in (BUILD_PS1, HARNESS_PS1):
        code = _ps_code(path)
        assert "$transient = @(Get-TransientFailures)" in code, (
            f"{path.name}: Get-TransientFailures must be materialised at the caller"
        )
        assert "$transient = Get-TransientFailures" not in code, (
            f"{path.name}: the unmaterialised call is still present"
        )
        assert "$transient.Count" in code, (
            f"{path.name}: the report must still test the count"
        )


def test_46m_every_collection_helper_is_materialised_at_every_caller() -> None:
    """The rule, applied uniformly rather than at the one site that happened to fire.

    Fixing only the reported call site would leave the same latent failure at every
    other caller, waiting for the run where that particular collection comes back
    empty. Every call to a collection-returning helper is wrapped at the caller.
    """
    names = "|".join(COLLECTION_HELPERS)
    call = re.compile(rf"(?<![\w-])({names})\b")
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        for number, line in enumerate(_ps_structural_lines(path), 1):
            if re.search(r"^\s*function\s+", line):
                continue
            for match in call.finditer(line):
                start = match.start()
                # Accept `@(Helper` -- the call is materialised right here.
                if line[max(0, start - 2) : start] == "@(":
                    continue
                problems.append(f"{path.name}:{number}: {line.strip()[:72]}")
    assert not problems, (
        "collection helper called without caller-side @(...):\n  " + "\n  ".join(problems)
    )


def test_46n_no_collection_result_reaches_count_unmaterialised() -> None:
    """The specific shape the strict-mode exception needs: assign bare, then .Count.

    Stated as its own test because it is the failure mode, not merely a style rule,
    and because it stays meaningful even if the helper inventory above changes.
    """
    names = "|".join(COLLECTION_HELPERS)
    bare = re.compile(rf"^\s*\$(\w+)\s*=\s*(?!@\()({names})\b")
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        lines = _ps_structural_lines(path)
        for number, line in enumerate(lines, 1):
            match = bare.match(line)
            if not match:
                continue
            variable = match.group(1)
            window = "\n".join(lines[number - 1 : number + 40])
            if re.search(rf"\${variable}\s*\.\s*Count\b", window):
                problems.append(
                    f"{path.name}:{number}: ${variable} = {match.group(2)} then ${variable}.Count"
                )
    assert not problems, (
        "unmaterialised collection reaching .Count (PropertyNotFoundException under "
        "Set-StrictMode):\n  " + "\n  ".join(problems)
    )


def test_46o_the_row_producer_emits_one_object_per_table_row() -> None:
    """The producer contract: 0 rows -> 0 objects, 1 row -> 1 row, N rows -> N rows.

    Two wrong shapes shipped before this one, and each looked right in isolation:

      `return $rows`   PowerShell enumerates a collection on output, so ONE row
                       goes out as its own cells and the caller sees N rows of one
                       cell each.
      `return ,$rows`  the unary comma emits the WHOLE jagged array as a SINGLE
                       object, so the caller's @(...) wraps that and lands one
                       level too deep: Count 1, element 0 the entire table. Every
                       `foreach ($row in ...)` then binds $row to the table.

    Neither is repairable at the caller, and neither is what `@(...)` is for --
    @(...) normalises 0/1/N, it does not fix nesting. So the producer emits each
    row itself, and this test pins that rather than the shape it replaced.
    """
    for helper in JAGGED_COLLECTION_HELPERS:
        body = _ps_function_body(HARNESS_PS1, helper)
        assert "return ,$rows" not in body, (
            f"{helper}: `return ,$rows` emits the whole table as one object"
        )
        assert "return $rows" not in body, (
            f"{helper}: `return $rows` unrolls a single row into its cells"
        )
        assert "$rows +=" not in body and "$rows = @()" not in body, (
            f"{helper}: rows must be emitted, not accumulated into a jagged array"
        )
        # A null DataBodyRange emits nothing at all.
        assert "if ($null -eq $body) { return }" in body, (
            f"{helper}: an empty body must emit zero objects, not an empty array"
        )
        # Each completed row goes out with non-enumerating row semantics.
        assert "Write-RowObject $line" in body, (
            f"{helper}: each row must be emitted through the shared row-emission idiom"
        )
        emitting = [
            line.strip()
            for line in body.splitlines()
            if "Write-RowObject" in line or re.search(r"(?<![\w-])return\b", line)
        ]
        assert emitting == ["if ($null -eq $body) { return }", "Write-RowObject $line"], (
            f"{helper}: unexpected output statements {emitting}"
        )

    # The shared idiom, and the reason it is shared.
    idiom = _ps_function_body(HARNESS_PS1, "Write-RowObject")
    assert "Write-Output -NoEnumerate $Row" in idiom, (
        "the row must be written without enumeration, or its cells are emitted "
        "as separate objects"
    )

    # A STREAMING producer has a new failure mode a collecting one did not: any
    # statement whose value is not consumed joins the output, and an extra object
    # between two rows would shift every later comparison by one. Every line in the
    # body must therefore be an assignment, a control construct, or one of the two
    # calls known to emit nothing.
    for helper in JAGGED_COLLECTION_HELPERS:
        leaks = []
        for line in _ps_function_body(HARNESS_PS1, helper).splitlines():
            statement = line.strip()
            if not statement or statement in ("{", "}", "} else {"):
                continue
            if statement.startswith("}"):
                statement = statement.lstrip("} ").strip()
                if not statement:
                    continue
            if re.match(
                r"(function|param|if|for|foreach|while|do|try|finally|catch|else|return)\b",
                statement,
            ):
                continue
            if "=" in statement:
                continue
            if re.match(r"(Release-Transient|Write-RowObject)\b", statement):
                continue
            leaks.append(statement[:70])
        assert not leaks, (
            f"{helper} streams its output, so an unconsumed statement would be "
            f"emitted between rows:\n  " + "\n  ".join(leaks)
        )

    # And the release helper it calls must itself be silent, for the same reason.
    release = _ps_function_body(LIFECYCLE_PS1, "Release-Transient")
    assert "$rec = Release-ComObjectSafe" in release, (
        "Release-Transient must consume its inner call's return value"
    )
    assert "$null = $script:transientFailures.Add(" in release, (
        "ArrayList.Add returns an index; discarding it keeps the helper silent"
    )

    # The flat helpers deliberately keep collecting and returning: one string
    # emitted is one string, which @(...) turns into a one-element array.
    for helper in ("Get-IdColumnValues", "Get-TableColumnNames"):
        body = _ps_function_body(HARNESS_PS1, helper)
        assert "return ," not in body, f"{helper} is flat and needs no unary comma"
        assert "Write-RowObject" not in body, f"{helper} emits values, not rows"


def test_46o1_no_caller_expects_the_whole_table_as_one_object() -> None:
    """The nested shape leaves a fingerprint at the caller: [0] used as the table.

    If a caller had been adjusted to the wrong producer -- `$body[0]` meaning the
    whole table, or `$body.Count -eq 1` standing for 'the table' -- fixing the
    producer would silently break it. Every consumer must read rows[n] = cells.
    """
    lines = _ps_structural_lines(HARNESS_PS1)
    table_vars = set()
    for line in lines:
        match = re.match(r"\s*\$(\w+)\s*=\s*@\(Get-TableBody\b", line)
        if match:
            table_vars.add(match.group(1))
    assert table_vars, "no Get-TableBody consumers found"

    problems = []
    for number, line in enumerate(lines, 1):
        for variable in table_vars:
            # `foreach ($x in $var[0])` would be iterating the table's first row as
            # if it were the table itself.
            if re.search(rf"foreach\s*\(\s*\$\w+\s+in\s+\${variable}\[0\]", line):
                problems.append(f"{number}: {line.strip()[:70]}")
    assert not problems, (
        "caller compensating for a nested producer shape:\n  " + "\n  ".join(problems)
    )


def test_46o2_the_preflight_proves_the_shape_before_excel_starts() -> None:
    """Linux cannot execute PowerShell; only the target can observe pipeline shape.

    Two wrong shapes got past source review, so the harness proves the contract
    itself, with fabricated arrays, no COM, and no Excel process -- and aborts the
    run if it fails, rather than comparing wrong shapes for twenty scenarios.
    """
    code = _ps(HARNESS_PS1)
    section = "# PRE. Collection shape, BEFORE Excel is started"
    assert section in code, "the harness has no collection-shape preflight"

    # It must run before anything starts Excel: the bootstrap is invoked in A.
    assert code.index(section) < code.index("# A. Stage-B build"), (
        "the preflight must run before the Stage-B bootstrap starts Excel"
    )
    assert code.index(section) < code.index("New-Object -ComObject Excel.Application")

    for check in (
        "zero rows: the caller collection is empty",
        "one row: the caller collection holds exactly one row",
        "one row: that row still has its three cells",
        "one row: the cell values are unchanged",
        "two rows: the caller collection holds exactly two rows",
        "two rows: each element is one row, boundaries preserved",
        "the whole collection is never emitted as one object",
    ):
        assert check in code, f"the preflight is missing: {check}"

    # A failed preflight must stop the run, not merely be recorded.
    preflight = code[code.index(section) : code.index("# Prepare a disposable copy")]
    assert "if (-not $preflightOk) {" in preflight and "exit 1" in preflight, (
        "a failed preflight must abort before Excel is started"
    )

    # ONE mechanism, exercised by both. A probe with its own copy of the idiom can
    # pass while the real reader is broken.
    probe = _ps_function_body(HARNESS_PS1, "Write-FabricatedRows")
    assert "Write-RowObject $line" in probe, (
        "the probe must exercise the same emission idiom as Get-TableBody"
    )
    assert "Write-Output" not in probe, (
        "the probe must not reimplement emission; it goes through Write-RowObject"
    )


# Two contracts that look alike in PowerShell and are not the same thing. Every
# helper that touches a collection belongs to exactly one of them.
#
#   ELEMENT PRODUCER  emits zero, one or many VALUES. The caller materialises with
#                     @(...). Get-TransientFailures, Get-TableBody, Write-RowObject.
#   CONTAINER FACTORY returns ONE object, which may be empty at birth. The caller
#                     keeps it and mutates it. New-Checklist, New-ReleaseLedger.
#
# Conflating them is what ended Gate-B run 2: `return (New-Object ArrayList)` reads
# as a factory but behaves as a producer, because an ArrayList is enumerable and an
# EMPTY enumerable emits zero pipeline objects.
CONTAINER_FACTORIES = ("New-Checklist", "New-ReleaseLedger")


def test_46q_the_checklist_factory_returns_one_mutable_arraylist() -> None:
    """`return (New-Object System.Collections.ArrayList)` returns nothing at all.

    An ArrayList is enumerable, and an empty enumerable emits ZERO pipeline objects,
    so `$list = New-Checklist` assigned $null and the first `$List.Add(...)` threw
    "You cannot call a method on a null-valued expression" -- in the preflight,
    before Excel started, on run 2. It is a factory, not a producer: it must hand
    back the object itself.
    """
    body = _ps_function_body(HARNESS_PS1, "New-Checklist")
    assert "Write-Output -NoEnumerate $list" in body, (
        "the factory must emit the ArrayList itself, not enumerate it"
    )
    assert "return (New-Object System.Collections.ArrayList)" not in body, (
        "returning the new ArrayList directly emits zero objects while it is empty"
    )
    assert not re.search(r"(?<![\w-])return\s+\$list\b", body), (
        "`return $list` enumerates the empty ArrayList and emits nothing"
    )
    assert "New-Object System.Collections.ArrayList" in body, (
        "callers rely on mutable .Add(); a plain array would not do"
    )


def test_46q1_every_checklist_call_site_receives_the_mutable_object() -> None:
    """21 call sites, and each must get the ArrayList, not a copy of its elements.

    `@(New-Checklist)` would satisfy the collection-materialisation rule and break
    every caller: @(...) yields an object[] of the ArrayList's elements, which has
    no .Add(). The two rules apply to different helpers and must not cross.
    """
    lines = _ps_structural_lines(HARNESS_PS1)
    sites = [n for n, line in enumerate(lines, 1) if re.search(r"=\s*New-Checklist\b", line)]
    assert len(sites) >= 21, f"only {len(sites)} checklist call sites found"
    wrapped = [
        f"{n}: {lines[n - 1].strip()[:60]}"
        for n in sites
        if re.search(r"=\s*@\(\s*New-Checklist", lines[n - 1])
    ]
    assert not wrapped, (
        "a checklist must not be materialised as a collection -- the caller needs "
        "the mutable object:\n  " + "\n  ".join(wrapped)
    )
    # And every one of them is used through .Add(), via Add-Check.
    add_check = _ps_function_body(HARNESS_PS1, "Add-Check")
    assert "$List.Add(" in add_check, "Add-Check must mutate the checklist in place"


def test_46q2_no_other_factory_returns_an_enumerable_empty_container() -> None:
    """The general rule, so the next factory of this shape is caught on Linux.

    A function that builds a container and hands it back must either emit it
    non-enumerated, or wrap it in something that is not enumerable -- a
    PSCustomObject, which is what New-ReleaseLedger does.
    """
    problems = []
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        for name in re.findall(r"^\s*function\s+([\w-]+)", _ps_code(path), re.MULTILINE):
            body = _ps_function_body(path, name)
            if "New-Object System.Collections.ArrayList" not in body:
                continue
            returns_container = re.search(
                r"(?<![\w-])return\s+(\$\w+\s*$|\(New-Object System\.Collections\.ArrayList\))",
                body,
                re.MULTILINE,
            )
            if not returns_container:
                continue  # it returns a string, a count, or nothing
            if "Write-Output -NoEnumerate" in body or "[pscustomobject]" in body:
                continue
            problems.append(f"{path.name}: {name} returns an enumerable container")
    assert not problems, (
        "container factory that emits nothing when empty:\n  " + "\n  ".join(problems)
    )
    # The two known factories, classified explicitly so the distinction is recorded.
    checklist = _ps_function_body(HARNESS_PS1, "New-Checklist")
    ledger = _ps_function_body(LIFECYCLE_PS1, "New-ReleaseLedger")
    assert "Write-Output -NoEnumerate" in checklist
    assert "[pscustomobject]@{" in ledger, (
        "New-ReleaseLedger is safe because a PSCustomObject is not enumerable"
    )
    assert "Write-Output" not in ledger, "it is already scalar and should stay so"


def test_46q3_the_factory_probe_runs_before_any_checklist_is_used() -> None:
    """Test infrastructure must not rest on an untested prerequisite.

    The row-shape preflight builds its findings in a checklist, so it cannot also
    be what proves the checklist factory works: when the factory returned $null,
    the first Add-Check threw before a single row-shape check had run.
    """
    code = _ps(HARNESS_PS1)
    marker = "# PRE0. Checklist factory, BEFORE anything that uses a checklist"
    assert marker in code, "there is no checklist-factory prerequisite probe"
    probe_at = code.index(marker)

    # Before the row-shape preflight, and before the first Add-Check anywhere.
    assert probe_at < code.index("# PRE. Collection shape"), (
        "the factory probe must precede the preflight that uses a checklist"
    )
    first_add_check = _ps_code(HARNESS_PS1).index("Add-Check $list")
    assert _ps_code(HARNESS_PS1).index("$probeChecklist = New-Checklist") < first_add_check
    assert probe_at < code.index("New-Object -ComObject Excel.Application")

    probe = code[probe_at : code.index("# PRE. Collection shape")]
    # It must not use the machinery it is testing. Checked against CODE, since the
    # probe's own commentary names Add-Check when explaining why it avoids it.
    executable = _ps_code(HARNESS_PS1)
    probe_code = executable[
        executable.index("$probeChecklist = New-Checklist") : executable.index("$list = New-Checklist")
    ]
    assert "Add-Check" not in probe_code, (
        "the probe cannot use Add-Check: that is the thing that fails when the "
        "factory is broken"
    )
    for proof in (
        "$null -eq $probeChecklist",                       # non-null
        "-is [System.Collections.ArrayList]",              # correct type
        "$probeChecklist.Add('sentinel')",                 # .Add() succeeds
        "$probeChecklist.Count -ne 1",                     # Count becomes 1
        "$probeChecklist[0] -ne 'sentinel'",               # the value survives
        "$probeChecklist.Clear()",                         # left clean for real use
    ):
        assert proof in probe, f"the factory probe is missing: {proof}"
    assert "exit 1" in probe, "a broken factory must abort before Excel is started"
    assert "Add-Result 'PRE0'" in probe


def test_46p_scalar_helpers_were_not_swept_up_in_the_rewrite() -> None:
    """The instruction was to materialise at collection call sites, not to rewrite
    every helper. A scalar helper wrapped in @(...) would quietly turn a string into
    a one-element array and break every comparison against it."""
    code = _ps_code(HARNESS_PS1) + _ps_code(BUILD_PS1) + _ps_code(LIFECYCLE_PS1)
    for scalar in ("Get-NamedValue", "Get-TrustAccessGuidance", "Get-TableRowCount"):
        assert f"@({scalar}" not in code, (
            f"{scalar} returns a scalar and must not be materialised as a collection"
        )
    # Get-TrustAccessGuidance builds a list but joins it into one string, which is
    # why it is correctly absent from COLLECTION_HELPERS.
    guidance = _ps_function_body(LIFECYCLE_PS1, "Get-TrustAccessGuidance")
    assert "-join" in guidance, (
        "Get-TrustAccessGuidance must still return a single joined string"
    )
    assert "Get-TrustAccessGuidance" not in COLLECTION_HELPERS


_PS_KEYWORDS = (
    "try", "catch", "finally", "if", "elseif", "else", "foreach", "for", "while",
    "do", "switch", "function", "filter", "trap", "param", "process", "begin", "end",
)
# Braces are their OWN tokens and can never be swallowed by an adjacent run of
# non-space characters: `[pscustomobject]@{` must yield a `{`, not one opaque blob.
_PS_TOKEN = re.compile(
    r"[{}]|(?<![\w-])(" + "|".join(_PS_KEYWORDS) + r")(?![\w-])|[^\s{}]+",
    re.IGNORECASE,
)


def _ps_block_frames(path: Path) -> tuple[list[tuple[int, str, str, str]], list[str]]:
    """Walk the script's blocks, tracking WHAT each brace opened.

    Returns (problems, unclosed). A problem is a `catch` or `finally` that does not
    follow the closing brace of the construct it must attach to.

    This is deliberately a lightweight keyword-aware scanner, not a PowerShell
    grammar: it answers one question -- does each catch/finally attach to its try?
    """
    problems: list[tuple[int, str, str, str]] = []
    stack: list[str] = []
    pending: str | None = None
    just_closed: str | None = None
    for number, line in enumerate(_ps_structural_lines(path), 1):
        for match in _PS_TOKEN.finditer(line):
            token = match.group(0)
            lowered = token.lower()
            if token == "{":
                # `@{` is a HASHTABLE LITERAL, not a block, and must not consume the
                # keyword a following body brace is waiting for. In
                #     foreach ($p in @( @{...}, @{...} )) { ... }
                # the literal would otherwise steal the `foreach` label.
                if match.start() > 0 and line[match.start() - 1] == "@":
                    stack.append("hashtable")
                else:
                    stack.append(pending or "block")
                    pending = None
                just_closed = None
            elif token == "}":
                just_closed = stack.pop() if stack else "UNDERFLOW"
                pending = None
            elif lowered in ("catch", "finally"):
                allowed = ("try",) if lowered == "catch" else ("try", "catch")
                if just_closed not in allowed:
                    problems.append((number, lowered, just_closed or "nothing", line.strip()[:60]))
                pending = lowered
                just_closed = None
            elif lowered in _PS_KEYWORDS:
                pending = lowered
                just_closed = None
            else:
                just_closed = None
    return problems, stack


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


def test_46k1_every_catch_and_finally_attaches_to_its_try() -> None:
    """BALANCED BRACES DO NOT IMPLY A VALID TRY/CATCH RELATIONSHIP.

    Wrapping seven scenario bodies in a prerequisite `if` produced

        try {
            if (Test-CleanStructure ...) {
                ...
            } catch {          <- this `}` closes the IF; catch attaches to nothing
                ...
        }
        }

    which balances perfectly and is a parse error. PowerShell parses the whole
    script before executing anything, so it would have stopped the harness before
    PRE0 -- and test_46k passed on it.
    """
    for path in (LIFECYCLE_PS1, BUILD_PS1, HARNESS_PS1):
        problems, unclosed = _ps_block_frames(path)
        rendered = [
            f"{path.name}:{number}: `{keyword}` follows a closed `{closed}`, not a try"
            f"  |  {text}"
            for number, keyword, closed, text in problems
        ]
        assert not rendered, "misattached catch/finally:\n  " + "\n  ".join(rendered)
        assert not unclosed, f"{path.name}: blocks left open: {unclosed}"


def test_46k2_every_guarded_scenario_has_the_required_shape() -> None:
    """The seven scenarios wrapped in the clean-structure guard, checked by name.

    Required, exactly:

        try {
            if (Test-CleanStructure -ExcelApp $excel -ScenarioId '<ID>') {
                ...
            }
        } catch {
            Add-Result '<ID>' ... 'FAIL' (Format-Err $_)
        }

    The catch belongs to the try, never to the if; a contaminated prerequisite is
    a SKIP that Test-CleanStructure has already emitted, never a FAIL.
    """
    lines = _ps_structural_lines(HARNESS_PS1)
    raw = _ps_code(HARNESS_PS1).splitlines()
    # The scenario id lives inside quotes, so it is read from the comment-stripped
    # source rather than the string-blanked one. Both keep line numbers aligned.
    guards = [
        (number, match.group(1))
        for number, line in enumerate(raw, 1)
        if (match := re.search(r"Test-CleanStructure -ExcelApp \$excel -ScenarioId '(\w+)'", line))
    ]
    assert {sid for _, sid in guards} == {"P", "Q", "R", "S", "T", "U", "W"}, guards

    for number, scenario in guards:
        # 1. exactly one enclosing try, opened before the guard.
        opener = number - 1
        while opener > 0 and lines[opener - 1].strip() != "try {":
            opener -= 1
        assert opener > 0, f"{scenario}: no enclosing `try {{` found"
        assert number - opener < 12, f"{scenario}: the guard is not directly inside its try"

        # Walk from the try, tracking what each brace opened.
        depth, frames, index = 0, [], opener - 1
        guard_closed_at = try_closed_at = catch_at = None
        while index < len(lines):
            line = lines[index]
            for match in _PS_TOKEN.finditer(line):
                token = match.group(0)
                if token == "{":
                    if match.start() > 0 and line[match.start() - 1] == "@":
                        frames.append("hashtable")
                    else:
                        frames.append("block")
                    depth += 1
                elif token == "}":
                    depth -= 1
                    frames.pop()
                    # 2. the guard `if` closes first, at depth 1 inside the try.
                    if depth == 1 and guard_closed_at is None and index + 1 > number:
                        guard_closed_at = index + 1
                    # 3. then the try itself closes.
                    elif depth == 0 and try_closed_at is None:
                        try_closed_at = index + 1
                elif token.lower() == "catch" and try_closed_at is not None and catch_at is None:
                    # 4. and only THEN does its catch begin.
                    catch_at = index + 1
            if catch_at is not None:
                break
            index += 1

        assert guard_closed_at is not None, f"{scenario}: the guard `if` is never closed"
        assert try_closed_at is not None, f"{scenario}: the enclosing `try` is never closed"
        assert catch_at is not None, f"{scenario}: the try has no catch"
        assert guard_closed_at < try_closed_at <= catch_at, (
            f"{scenario}: guard closes at {guard_closed_at}, try at {try_closed_at}, "
            f"catch at {catch_at} -- the catch must follow the TRY's closing brace"
        )

        # 5. no catch is attached immediately after the inner if block.
        assert "catch" not in lines[guard_closed_at - 1].lower(), (
            f"{scenario}:{guard_closed_at}: a catch is attached to the guard `if`"
        )

        # 6. exactly one catch for this try, and it reports a FAIL for this scenario.
        body = "\n".join(raw[opener - 1 : catch_at + 3])
        assert body.count("} catch {") == 1, f"{scenario}: more than one catch on its try"
        assert f"Add-Result '{scenario}'" in "\n".join(raw[catch_at - 1 : catch_at + 2]), (
            f"{scenario}: the catch does not report a FAIL for its own scenario"
        )
        # Contamination stays a SKIP: the guard emits it, the scenario body simply
        # does not run.
        guard_body = _ps_function_body(HARNESS_PS1, "Test-CleanStructure")
        assert "'SKIP'" in guard_body and "'FAIL'" not in guard_body


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
