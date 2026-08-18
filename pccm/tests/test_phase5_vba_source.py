#!/usr/bin/env python3
"""PCCM Phase 5 Gate A Step 4: STATIC tests over the pure VBA numerical kernel.

NO VBA IS EXECUTED HERE, AND NONE CAN BE. VBA has no interpreter on Linux, so
every assertion in this file is a statement about SOURCE TEXT: which procedures
exist, which constructs appear in executable code, which module owns which
responsibility, and which public names the kernel exposes.

Nothing in this file establishes that the VBA produces a correct number, that
its arithmetic behaves as the reference implementation does, or that any
fingerprint parity has been observed. Those are Gate B's claims, on real Excel
on Windows, and no test name, docstring or message here may suggest otherwise.

What this file DOES establish:

  * the pure-numerical boundary holds - no workbook, no Excel object, no
    randomness reaches the kernel
  * error handling is confined to the four arithmetic primitives, in the locked
    shape
  * each module owns its declared responsibility and nothing else
  * the public API of each module is exactly the whitelisted surface, so the
    kernel cannot grow an accidental entry point
  * the constructs the accepted plan forbids are absent, and the constructs it
    requires are present
  * twelve deliberately planted defects are each REJECTED by the sweep that
    exists to catch them - so a passing suite means the sweeps still work

Runs standalone or under pytest.
"""

from __future__ import annotations

import hashlib
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder.vba_source import (  # noqa: E402
    VbaModule,
    load_modules,
    logical_statements,
    strip_comments,
    strip_strings,
)

SRC_VBA = PCCM_ROOT / "src" / "vba"

KERNEL_MODULES = ("modCalcFactors", "modCalcAnalytical", "modCalcFingerprint")

# The Step-5 resolution layer. It is the ONE module allowed workbook access, and
# it is deliberately not in KERNEL_MODULES: every worksheet-free sweep below
# runs over the kernel and must keep running over exactly the kernel.
STEP5_MODULE = "modCalcResolve"

# The Step-6 numerical prerequisite checker. Like the resolver it is outside
# KERNEL_MODULES, so every sweep below keeps running over exactly the kernel.
STEP6_MODULE = "modCalcCheck"

PHASE4_MODULES = (
    "modWorkbook", "modAppState", "modTimeline", "modDrivers",
    "modProfiling", "modInflation", "modStructuralCheck",
)

# ---------------------------------------------------------------------------
# The pure-numerical boundary
# ---------------------------------------------------------------------------
# Every one of these is a way for the workbook, the host application or a random
# number generator to reach the kernel. The kernel is handed numbers and hands
# back numbers; a later resolver layer is what touches Excel.
BOUNDARY_TOKENS = (
    "Application.", "ThisWorkbook", "ActiveWorkbook", "Worksheets", "Worksheet",
    "Range", "Cells", "ListObjects", "ListObject", "Names(", "Evaluate",
    "WorksheetFunction", "modWorkbook.", "Rnd", "Randomize", "MRG32k3a",
    "NPV", "Percentile",
)

# Types a kernel parameter may have. Anything else - and in particular any Excel
# object - is refused.
# VARIANT IS NOT ON THIS LIST. The pure numerical modules are typed end to end:
# a container that can hold anything gives up the type checking that would catch
# a wrong numerical shape here rather than on Windows.
ALLOWED_PARAMETER_TYPES = {
    "Double", "Long", "Boolean", "String",
    # The locked carry types and the module-local records.
    "DriverFactors", "YearFactors", "DriverAudit", "AnalyticalTotals",
    "AnnualRow", "ReconciliationMagnitudes", "IdentityCheck",
    # The exact kernel's private record: (sign, base-2^24 limbs, binary shift).
    "ExactNumber",
}

# ---------------------------------------------------------------------------
# The public API whitelist - EXACT, in both directions
# ---------------------------------------------------------------------------
# "At least the required names exist" is not enough: a kernel that quietly grew a
# public helper would have grown an entry point nobody reviewed. Each name below
# is public because something outside its module calls it.
FACTORS_PUBLIC = {
    # The range predicate and the four arithmetic primitives.
    "IsUsableDouble", "SafeAdd", "SafeSubtract", "SafeMultiply", "SafeDivide",
    "SafeAccumulate",
    # The two-tier rescues.
    "SafeSignedSum", "SafeProduct", "ExactSumOfProducts",
    # PUBLIC BECAUSE modCalcAnalytical CALLS IT: the convex-statistic exact tier
    # needs (sum of terms) / divisor for divisor in {2, 3, 6}, and that rounding
    # is the exact kernel's business, not the analytical layer's. Duplicating it
    # would be a second rounding rule.
    "ExactQuotientOfSum",
    # The factor series and the C2 materialization boundary.
    "BuildInflationFactors", "BuildDiscountFactors", "BuildKnom", "BuildKpv",
    # C1 conditioning.
    "ConditioningScaledMagnitude",
    # PUBLIC BECAUSE modCalcAnalytical CALLS IT: an annual contribution with no
    # Double of its own still has to be conditioned, and only the exact kernel
    # can fold the coefficient into that factor expression.
    "ConditioningScaledProduct",
    "IdentityAllowance",
}

ANALYTICAL_PUBLIC = {
    "TriangularMean", "PertMean", "UniformMean", "DeterministicCentral", "ExpectedRisk",
    "BuildDriverAudit", "AccumulateTotals", "BuildAnnualSeries", "Reconcile",
    # Later orchestration consumes the reconciliation result, so the verdict over
    # a check array is part of the reviewed surface even though nothing calls it
    # across a module boundary yet.
    "AllIdentitiesHold",
}

FINGERPRINT_PUBLIC = {
    "CalcFpUtf16Length", "CalcFpNormaliseCodeUnit", "CalcFpCanonicalText",
    "CalcFpCanonicalNumber", "CalcFpCanonicalInteger",
    "CalcFpReduceDouble", "CalcFpDigestStream", "CalcFpBuildCostRecord",
    "CalcFpBuildRiskRecord", "CalcFpBuildFingerprint",
}

# Public WITHOUT a current cross-module caller, each for a stated reason. Every
# other Public name must have one, and test_71 proves it by scanning references.
PUBLIC_WITHOUT_CROSS_MODULE_CALLER = {
    # The Gate-B diagnostic surface: the locked helper vectors are compared one
    # by one against the reference implementation on Windows, so each stays
    # reachable even though the production path goes through the builders.
    "CalcFpUtf16Length", "CalcFpNormaliseCodeUnit", "CalcFpCanonicalText",
    "CalcFpCanonicalNumber", "CalcFpCanonicalInteger", "CalcFpReduceDouble",
    "CalcFpDigestStream",
    # Consumed by later orchestration, not by another module today.
    "AllIdentitiesHold",
    # The primitives and rescues the resolver layer will call directly.
    "ExpectedRisk", "IsUsableDouble", "SafeAdd", "SafeSubtract", "SafeMultiply", "SafeDivide",
    "SafeAccumulate", "SafeSignedSum", "BuildInflationFactors",
    "BuildDiscountFactors", "BuildKnom", "BuildKpv", "IdentityAllowance",
    "TriangularMean", "PertMean", "UniformMean", "DeterministicCentral",
    "BuildDriverAudit", "AccumulateTotals", "BuildAnnualSeries", "Reconcile",
    "CalcFpBuildCostRecord", "CalcFpBuildRiskRecord", "CalcFpBuildFingerprint",
}

PUBLIC_SURFACE = {
    "modCalcFactors": FACTORS_PUBLIC,
    "modCalcAnalytical": ANALYTICAL_PUBLIC,
    "modCalcFingerprint": FINGERPRINT_PUBLIC,
}

# The four procedures that may install an error handler. Each scopes it to ONE
# arithmetic expression; everywhere else, a failure is a returned False.
HANDLER_OWNERS = {"SafeAdd", "SafeSubtract", "SafeMultiply", "SafeDivide"}

# The Phase-4 VBA, byte for byte as Step 3 left it. Step 4 adds modules; it does
# not touch one.
PHASE4_SHA256 = {
    "modWorkbook": "9cfa8f130c5bcdee783948654c969d4b0d6589fe7059c126f88c7676ca5405bf",
    "modAppState": "ef0b5c64a7a3b5aeeef5ef0797cd160071a7eda6a7d8cef9cb98301f1504672f",
    "modTimeline": "4a4f24d17b65bcbc0e46b1a74213b6a02eab6ab492b1788476d66eb7807b9e3f",
    "modDrivers": "8f947a4cc473b76161c867f99daf5fbb4af670b909cca0387165b079c102af48",
    "modProfiling": "0312858d7d817d20a99877f8be52ca0f7cf5b0bbb9aa9770367ed11138d9d7ca",
    "modInflation": "08db32807d495c22e6067350291c21a9a277884de5e5064555612f6bb991118c",
    "modStructuralCheck": "1798c56a459c9e35c581871248815841b28a3c88a62a931a68afe5d71853ed54",
}


# ===========================================================================
# fixtures and sweeps
# ===========================================================================
def _modules() -> dict[str, VbaModule]:
    return {m.name: m for m in load_modules([SRC_VBA])}


def _kernel() -> dict[str, VbaModule]:
    modules = _modules()
    return {name: modules[name] for name in KERNEL_MODULES}


def _synthetic(name: str, body: str) -> VbaModule:
    """A module built from text, for the negative controls.

    Nothing is written to disk and nothing is executed: the sweeps read text, so a
    planted defect can be handed to them as text.
    """
    return VbaModule(name=name, path=SRC_VBA / f"{name}.bas", raw=body)


def boundary_hits(module: VbaModule) -> list[str]:
    """Boundary tokens appearing in EXECUTABLE code.

    Comments and string literals are removed first, so a comment explaining why
    the kernel never touches a Worksheet is not read as touching one.
    """
    code = module.code
    return sorted({token for token in BOUNDARY_TOKENS if token.lower() in code.lower()})


def parameter_types(module: VbaModule) -> set[str]:
    """Every declared parameter type in the module's procedure signatures."""
    found: set[str] = set()
    for _, statement in logical_statements(module.code_without_string_removal):
        if not re.match(r"^\s*(Public |Private |Friend )?(Static )?(Sub|Function)\s", statement):
            continue
        inner = statement[statement.find("(") + 1 : statement.rfind(")")]
        for part in inner.split(","):
            match = re.search(r"\bAs\s+([A-Za-z_]\w*)", part)
            if match:
                found.add(match.group(1))
    return found


def handler_owners(module: VbaModule) -> set[str]:
    """Procedures containing `On Error`, by name."""
    owners: set[str] = set()
    current = ""
    for line in module.code.splitlines():
        match = re.match(
            r"^\s*(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function)\s+(\w+)", line
        )
        if match:
            current = match.group(1)
        if re.search(r"\bOn\s+Error\b", line):
            owners.add(current)
    return owners


def uses_native_integer_division(code: str) -> bool:
    """`Mod` or `\\` as an operator - both Long-typed in VBA, both refused."""
    return bool(re.search(r"(?<![\w.])Mod(?![\w])", code, re.IGNORECASE)) or "\\" in code


def wide_type_hits(code: str) -> list[str]:
    """Types and helpers the exact kernel may not use."""
    return sorted(
        keyword
        for keyword in ("Currency", "Decimal", "CDec", "CCur", "Eval", "WorksheetFunction")
        if re.search(rf"(?<![\w.]){keyword}(?![\w])", code, re.IGNORECASE)
    )


# ===========================================================================
# 1. the modules themselves
# ===========================================================================
def test_01_the_three_kernel_modules_exist_and_declare_themselves() -> None:
    for name, module in _kernel().items():
        lines = module.raw.splitlines()
        assert lines[0] == f'Attribute VB_Name = "{name}"', (
            f"{name} must open with its own VB_Name attribute"
        )
        assert lines[1] == "Option Explicit", f"{name} must declare Option Explicit"


def test_02_step_4_added_exactly_three_modules_and_no_fourth() -> None:
    """The accepted Step-4 module split is deliberate.

    A modCalcMath, modCalcTypes or modCalcExact would be a fourth NUMERICAL
    module that no review accepted, and the split would stop meaning anything.
    Steps 5 and 6 add exactly one further module each - the resolver and the
    checker - and the inventory is asserted in both directions so no step can
    grow another.
    """
    on_disk = set(_modules())
    assert on_disk == (set(PHASE4_MODULES) | set(KERNEL_MODULES)
                       | {STEP5_MODULE, STEP6_MODULE}), (
        f"unexpected hand-written module inventory: {sorted(on_disk)}"
    )
    assert set(KERNEL_MODULES) == {
        "modCalcFactors", "modCalcAnalytical", "modCalcFingerprint"
    }


def test_03_no_phase4_vba_source_file_changed() -> None:
    """Byte-for-byte, against the digests recorded when Step 3 was accepted."""
    for name, digest in PHASE4_SHA256.items():
        actual = hashlib.sha256((SRC_VBA / f"{name}.bas").read_bytes()).hexdigest()
        assert actual == digest, f"{name}.bas changed; Phase-4 VBA is frozen"


# ===========================================================================
# 2. the pure-numerical boundary
# ===========================================================================
def test_04_no_kernel_module_touches_the_workbook_or_the_host() -> None:
    for name, module in _kernel().items():
        assert boundary_hits(module) == [], (
            f"{name} reaches outside the numerical boundary: {boundary_hits(module)}"
        )


def test_05_the_boundary_sweep_reads_code_and_not_commentary() -> None:
    """A comment saying the kernel never touches a Worksheet is not a violation.

    Stated as its own test because the sweep is only trustworthy if it can tell
    the two apart, and every kernel module does discuss what it does not do.
    """
    planted = _synthetic(
        "modProbe",
        'Attribute VB_Name = "modProbe"\nOption Explicit\n'
        "' This module never calls Application.Calculate or reads a Worksheet.\n"
        'Public Function Describe() As String\n'
        '    Describe = "Worksheets are not touched"\n'
        "End Function\n",
    )
    assert boundary_hits(planted) == []


def test_06_no_excel_object_appears_in_a_parameter_type() -> None:
    for name, module in _kernel().items():
        unexpected = parameter_types(module) - ALLOWED_PARAMETER_TYPES
        assert not unexpected, f"{name} declares parameter type(s) {sorted(unexpected)}"


def test_07_no_kernel_procedure_is_a_pccm_endpoint() -> None:
    """Step 4 is source only. A callable macro would be an endpoint nobody asked for."""
    for name, module in _kernel().items():
        offenders = [p for p in module.procedures if p.startswith("PCCM_")]
        assert not offenders, f"{name} declares {offenders}"


def test_08_the_deferred_phase_6_surface_does_not_exist_yet() -> None:
    """Each name leaves this list at the step that implements it.

    `modCalcResolve` left at Step 5 and `modCalcCheck` at Step 6. Everything
    still on it is ahead: the reporter and the five Phase-5 status accessors.
    """
    deferred = (
        "modCalcReport",
        "PCCM_Calculate", "PCCM_CalculationStatus", "PCCM_CalculationAttemptResult",
        "PCCM_CalculationAttemptDetail", "PCCM_CalculationFingerprint",
        "PCCM_CurrentInputFingerprint",
    )
    # EXECUTABLE code, not commentary: the resolver legitimately names the
    # checker when saying which prerequisites it deliberately leaves to it, and
    # a sentence about a later step is not an implementation of one.
    modules = _modules()
    executable = "\n".join(m.code for m in modules.values())
    declared = {p for m in modules.values() for p in m.procedures}
    for name in deferred:
        assert name not in executable, f"{name} is referenced in code; it belongs to a later step"
        assert name not in declared, f"{name} is declared; it belongs to a later step"
        assert name not in modules, f"{name} exists as a module; it belongs to a later step"


def test_09_no_change_handler_exists_anywhere() -> None:
    everything = "\n".join(m.code for m in _modules().values())
    for handler in ("Worksheet_Change", "Workbook_SheetChange", "Worksheet_Calculate"):
        assert handler not in everything


# ===========================================================================
# 3. error handling
# ===========================================================================
def test_10_no_kernel_module_uses_on_error_resume_next() -> None:
    for name, module in _kernel().items():
        assert "On Error Resume Next" not in module.code, (
            f"{name} suppresses errors; a suppressed error is a silent wrong answer"
        )


def test_11_only_the_four_arithmetic_primitives_install_a_handler() -> None:
    factors = _kernel()["modCalcFactors"]
    assert handler_owners(factors) == HANDLER_OWNERS
    for name in ("modCalcAnalytical", "modCalcFingerprint"):
        assert handler_owners(_kernel()[name]) == set(), (
            f"{name} installs an error handler; failure there is a returned False"
        )


def test_12_each_handler_is_scoped_to_one_arithmetic_expression() -> None:
    """The locked shape, asserted line by line.

    `On Error GoTo ArithmeticFailure`, ONE expression, `On Error GoTo 0`, then the
    post-checks. A handler that stayed armed across the range checks would catch a
    defect in the checking, not in the arithmetic.
    """
    lines = _kernel()["modCalcFactors"].code.splitlines()
    armed = [i for i, line in enumerate(lines) if line.strip() == "On Error GoTo ArithmeticFailure"]
    assert len(armed) == len(HANDLER_OWNERS)
    for index in armed:
        expression = lines[index + 1].strip()
        assert re.match(r"^tmp = a [-+*/] b$", expression), (
            f"a handler guards {expression!r}, which is not a single arithmetic expression"
        )
        assert lines[index + 2].strip() == "On Error GoTo 0", (
            "the handler must be disarmed immediately after the guarded expression"
        )


def test_13_every_handler_label_disarms_and_returns_false() -> None:
    lines = _kernel()["modCalcFactors"].code.splitlines()
    labels = [i for i, line in enumerate(lines) if line.strip() == "ArithmeticFailure:"]
    assert len(labels) == len(HANDLER_OWNERS)
    for index in labels:
        assert lines[index + 1].strip() == "On Error GoTo 0"
        assert re.match(r"^Safe\w+ = False$", lines[index + 2].strip()), (
            "the failure path must leave the caller's ByRef result untouched and "
            "return False"
        )


def test_14_safe_accumulate_delegates_to_safe_add() -> None:
    body = _procedure_body(_kernel()["modCalcFactors"], "SafeAccumulate")
    assert "SafeAdd(" in body, "SafeAccumulate must not add on its own"


def _procedure_body(module: VbaModule, name: str) -> str:
    lines = module.code.splitlines()
    start = next(
        i for i, line in enumerate(lines)
        if re.match(rf"^\s*(Public |Private )?(Static )?(Sub|Function)\s+{name}\b", line)
    )
    end = next(i for i in range(start + 1, len(lines)) if re.match(r"^End (Sub|Function)", lines[i]))
    return "\n".join(lines[start:end])


# ===========================================================================
# 4. the exact kernel
# ===========================================================================
def test_15_the_limb_base_is_the_locked_one() -> None:
    raw = _kernel()["modCalcFactors"].raw
    assert "Private Const LIMB_BITS As Long = 24" in raw
    assert "Private Const LIMB_BASE As Double = 16777216#" in raw


def test_16_the_exact_kernel_uses_no_wide_type_and_no_integer_operator() -> None:
    for name, module in _kernel().items():
        assert wide_type_hits(module.code) == [], (
            f"{name} uses a type or helper the exact kernel forbids: "
            f"{wide_type_hits(module.code)}"
        )
        assert not uses_native_integer_division(module.code), (
            f"{name} uses Mod or \\, both of which VBA evaluates as Long"
        )


def test_17_the_kernel_never_narrows_a_wide_significand() -> None:
    """`CLng` is permitted only where the value is provably small.

    The single use is a hex digit, which is 0 to 15. Narrowing an accumulator or a
    limb product would reintroduce exactly the Long ceiling the kernel exists to
    avoid.
    """
    uses: list[str] = []
    for name, module in _kernel().items():
        for _, statement in logical_statements(module.code):
            for match in re.finditer(r"CLng\(([^)]*)\)", statement):
                uses.append(f"{name}: {match.group(1).strip()}")
    assert uses == ["modCalcFingerprint: digit"], f"unexpected CLng narrowing: {uses}"


def test_18_the_two_tier_rescues_are_the_exact_kernel_and_not_the_superseded_heuristics() -> None:
    """Tier 2 is exact, not a reordering.

    The rejected designs were a rounded positive-minus-negative cancellation and a
    magnitude-balanced product ordering. Both were shown to accept a wrong answer,
    so tier 2 must reach the exact routines and nothing else.
    """
    factors = _kernel()["modCalcFactors"]
    for name, exact_call in (("SafeSignedSum", "ExactSumOf("), ("SafeProduct", "ExactProductOf(")):
        body = _procedure_body(factors, name)
        assert exact_call in body, f"{name} does not reach the exact kernel"
        assert "RoundExact(" in body, f"{name} must round the exact value exactly once"


def test_19_the_factor_series_are_iterative_and_never_a_power() -> None:
    """`(1+r)^(t-1)` can overflow where the factor is representable, and it cannot
    say which year failed."""
    factors = _kernel()["modCalcFactors"]
    for name in ("BuildInflationFactors", "BuildDiscountFactors"):
        body = _procedure_body(factors, name)
        assert "^" not in body, f"{name} raises a power instead of iterating"
        assert "For " in body, f"{name} must build its series year by year"


def test_20_knom_and_kpv_exclude_quantity_and_probability() -> None:
    """Probability is replaced by a Bernoulli draw in Monte Carlo and Quantity is a
    per-driver multiplier; folding either in would double-count it later."""
    factors = _kernel()["modCalcFactors"]
    for name in ("BuildKnom", "BuildKpv", "BuildFactor"):
        body = _procedure_body(factors, name)
        for forbidden in ("Quantity", "Probability"):
            assert forbidden not in body, f"{name} references {forbidden}"


# ===========================================================================
# 5. tolerances and conditioning
# ===========================================================================
TOLERANCE_CONSTANTS = (
    "TOL_PROFILING_SUM_ABSOLUTE", "TOL_IDENTITY_ABSOLUTE_FLOOR",
    "TOL_IDENTITY_RELATIVE_COEFFICIENT", "TOL_CONDITIONING_SCALE_FLOOR",
)


def test_21_tolerances_are_read_from_the_generated_constants() -> None:
    analytical = _kernel()["modCalcAnalytical"].code
    for constant in TOLERANCE_CONSTANTS:
        assert constant in analytical, f"{constant} is never used"


def test_22_no_tolerance_number_is_restated_as_a_production_literal() -> None:
    """A tolerance written here would be a second authority.

    The numbers live in spec/calc_contract.yaml and reach the VBA only through the
    generated module.
    """
    for name, module in _kernel().items():
        for literal in ("1e-9", "1E-9", "1e-6", "1E-6", "1e-12", "1E-12"):
            assert literal not in module.code, f"{name} restates the tolerance {literal}"


# ===========================================================================
# 6. the analytical layer
# ===========================================================================
def test_23_the_five_measures_are_five_independent_passes() -> None:
    """B is NOT C - A and E is NOT C + D.

    Those are the reconciliation identities, and an identity computed by
    definition checks nothing. Each measure keeps its own contribution list.
    """
    body = _procedure_body(_kernel()["modCalcAnalytical"], "AccumulateTotals")
    for derived in (
        r"BNom\s*=.*CNom", r"BPv\s*=.*CPv",
        r"ENom\s*=.*CNom.*DNom", r"EPv\s*=.*CPv.*DPv",
    ):
        assert not re.search(derived, body), f"a measure is derived, not accumulated: {derived}"
    for terms in ("aNomTerms", "bNomTerms", "cNomTerms", "dNomTerms", "eNomTerms"):
        assert terms in body, f"{terms} is missing; each measure needs its own list"


def test_24_the_six_annual_series_are_six_separate_boundaries() -> None:
    body = _procedure_body(_kernel()["modCalcAnalytical"], "BuildAnnualSeries")
    calls = re.findall(r"(?<!Build)AnnualSeries\(", body)
    assert len(calls) == 6, f"expected six independent series, found {len(calls)}"
    for field in (
        "BaseCostNominal", "ExpectedRiskNominal", "TotalNominal",
        "BaseCostPv", "ExpectedRiskPv", "TotalPv",
    ):
        assert field in body, f"{field} is not produced by its own pass"
    assert not re.search(r"TotalNominal\s*=\s*.*BaseCostNominal", body), (
        "the annual total must be summed over its own contributions"
    )


def test_25_the_convex_statistics_have_three_tiers() -> None:
    analytical = _kernel()["modCalcAnalytical"]
    for name in ("TriangularMean", "PertMean", "UniformMean"):
        body = _procedure_body(analytical, name)
        assert "DegeneratePoint(" in body, f"{name} lacks the zero-uncertainty tier"
        assert "StableConvex(" in body, f"{name} lacks the ordinary staged tier"
        assert "ConvexFinish(" in body, f"{name} lacks the exact tier"
    finish = _procedure_body(analytical, "ConvexFinish")
    assert "ExactQuotientOfSum(" in finish
    assert "staged <> 0#" in finish, (
        "a staged zero can be an underflow hiding a small non-zero answer, so it "
        "must fall through to the exact numerator"
    )


def test_26_the_pert_numerator_is_four_copies_of_most_likely() -> None:
    """Forming `4 * ML` first is the avoidable overflow the stable form prevents."""
    body = _procedure_body(_kernel()["modCalcAnalytical"], "PertMean")
    copies = re.findall(r"numerator\(\d\) = mostLikely", body)
    assert len(copies) == 4, f"expected four copies of Most Likely, found {len(copies)}"
    assert "4 * mostLikely" not in body and "4# * mostLikely" not in body


def test_27_the_canonical_order_is_binary_string_comparison() -> None:
    """Ordinal UTF-16 code units, matching StrComp(..., vbBinaryCompare)."""
    for name in ("modCalcAnalytical", "modCalcFingerprint"):
        code = _kernel()[name].code
        assert "vbBinaryCompare" in code, f"{name} does not order by binary comparison"
        assert "vbTextCompare" not in code, f"{name} uses a text comparison"


def test_28_reconciliation_covers_i1_to_i5() -> None:
    raw = _kernel()["modCalcAnalytical"].raw
    for identity in ("I1 nominal", "I1 PV", "I2 nominal", "I2 PV",
                     "I3a", "I3b", "I3c", "I4a", "I4b", "I4c", "I5 profile sum"):
        assert identity in raw, f"{identity} is never produced"


# ===========================================================================
# 7. the fingerprint
# ===========================================================================
def test_29_every_required_fingerprint_helper_exists() -> None:
    available = set(_kernel()["modCalcFingerprint"].procedures)
    for name in (
        "CalcFpUtf16Length", "CalcFpNormaliseCodeUnit", "CalcFpCanonicalText",
        "CalcFpCanonicalNumber", "CalcFpCanonicalInteger", "CalcFpReduceDouble",
        "CalcFpDigestStream", "CalcFpBuildCostRecord", "CalcFpBuildRiskRecord",
        "CalcFpBuildFingerprint",
    ):
        assert name in available, f"{name} is missing"


def test_30_the_fingerprint_constants_are_projected_and_never_restated() -> None:
    """The moduli in particular. A literal here would be a second authority."""
    module = _kernel()["modCalcFingerprint"]
    for literal in ("2147483647", "2147483629", "131"):
        assert literal not in module.code, (
            f"the fingerprint module restates {literal}; it must use the FP_* constants"
        )
    for constant in ("FP_BASE", "FP_MOD_1", "FP_MOD_2", "FP_INIT_1", "FP_INIT_2",
                     "FP_STREAM_TAG", "FP_TAG_TEXT", "FP_TAG_NUMBER", "FP_TAG_INTEGER",
                     "FP_SECTION_1", "FP_SECTION_2", "FP_SECTION_3"):
        assert constant in module.code, f"{constant} is never used"


def test_31_the_reducer_is_the_locked_double_only_form() -> None:
    body = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpReduceDouble")
    statements = [text for _, text in logical_statements(body)]
    assert "x = h * FP_BASE + u" in statements
    assert "q = Fix(x / modulus)" in statements
    assert "r = x - q * modulus" in statements
    assert "If r >= modulus Then r = r - modulus" in statements
    assert "If r < 0# Then r = r + modulus" in statements
    assert not uses_native_integer_division(body)
    assert "CLng" not in body, "neither x nor q * modulus may be narrowed to a Long"


def test_32_the_digest_is_two_eight_digit_hex_accumulators() -> None:
    body = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpDigestStream")
    assert "h1 = FP_INIT_1" in body and "h2 = FP_INIT_2" in body
    assert "CalcFpReduceDouble(h1, unit, FP_MOD_1)" in body
    assert "CalcFpReduceDouble(h2, unit, FP_MOD_2)" in body
    assert "CalcFpHex8(h1) & CalcFpHex8(h2)" in body
    assert "Private Const FP_HEX_WIDTH As Long = 8" in _kernel()["modCalcFingerprint"].raw


def test_33_the_stream_is_walked_by_code_unit_with_signed_ascw_normalised() -> None:
    module = _kernel()["modCalcFingerprint"]
    digest = _procedure_body(module, "CalcFpDigestStream")
    assert "AscW(Mid$(stream, index, 1))" in digest, (
        "the stream must be walked one UTF-16 code unit at a time"
    )
    normalise = _procedure_body(module, "CalcFpNormaliseCodeUnit")
    assert "normalised = normalised + 65536" in normalise, (
        "AscW returns a signed Integer, so a unit above U+7FFF comes back negative"
    )
    length = _procedure_body(module, "CalcFpUtf16Length")
    assert "Len(text)" in length


def test_34_the_decimal_separator_is_an_argument_and_never_read_from_the_machine() -> None:
    module = _kernel()["modCalcFingerprint"]
    assert "Application.International" not in module.code
    for name in ("CalcFpCanonicalNumber", "CalcFpNumberField",
                 "CalcFpBuildCostRecord", "CalcFpBuildRiskRecord"):
        signature = _procedure_body(module, name).splitlines()
        joined = " ".join(line.strip().rstrip("_") for line in signature)
        assert "decimalSeparator As String" in joined, (
            f"{name} must be told which separator the host formatter produced"
        )


def test_35_the_separator_normalisation_is_positional_and_not_a_global_replace() -> None:
    """`E`, `+`, `-` and every digit already occur in scientific notation."""
    body = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpCanonicalNumber")
    assert "Replace" not in body, "a global replace would corrupt the exponent"
    assert "marker = CalcFpMarkerIndex(text)" in body, (
        "the marker must be located in the formatter's own output"
    )
    literal = _procedure_body_raw(_kernel()["modCalcFingerprint"], "CalcFpCanonicalNumber")
    assert 'result = Left$(text, marker - 1) & "." & Mid$(text, marker + 1)' in literal, (
        "exactly the marker position must be rewritten to a full stop"
    )


def host_marker_equality_gate(module: VbaModule, procedure: str) -> list[str]:
    """Statements that refuse when the host's own marker differs from the
    caller-supplied separator.

    Such a gate makes the locked dual injection unsatisfiable: on ONE host, both
    "." and "," must produce byte-identical output, and whichever character the
    formatter emits, a gate comparing against it refuses the other injection.
    """
    statements = [text for _, text in logical_statements(_procedure_body(module, procedure))]
    return [
        statement
        for statement in statements
        if re.search(r"Mid\$?\(\s*\w+\s*,\s*marker\s*,\s*1\s*\)\s*<>\s*decimalSeparator",
                     statement)
        or re.search(r"decimalSeparator\s*<>\s*Mid\$?\(", statement)
    ]


def test_35a_the_host_marker_is_never_compared_against_the_supplied_separator() -> None:
    """Gate A cannot run VBA, but it CAN prove the source does not make the two
    locked separator injections mutually exclusive.

    The accepted plan permits exactly this implementation: identify the mantissa
    marker from the formatter's own output and normalise that one position. The
    supplied separator therefore does not have to match what the host emitted,
    and requiring it to match is what would make one of the two injections
    impossible on any single machine.
    """
    module = _kernel()["modCalcFingerprint"]
    assert host_marker_equality_gate(module, "CalcFpCanonicalNumber") == [], (
        "the encoder refuses a separator that differs from the host's own marker; "
        "the locked pair of injections could then never both succeed"
    )
    statements = [
        text for _, text in
        logical_statements(_procedure_body(module, "CalcFpCanonicalNumber"))
    ]
    for statement in statements:
        assert not re.search(r"\bdecimalSeparator\b.*(<>|=).*\btext\b", statement), (
            "the separator must not be validated against the formatted text at all"
        )


def test_35b_the_separator_remains_a_validated_argument() -> None:
    """The public interface is unchanged and no machine state is consulted."""
    module = _kernel()["modCalcFingerprint"]
    assert "decimalSeparator As String" in _signature(module, "CalcFpCanonicalNumber")
    body = _procedure_body(module, "CalcFpCanonicalNumber")
    assert "If CalcFpUtf16Length(decimalSeparator) <> 1 Then Exit Function" in body, (
        "the separator is still validated as exactly one UTF-16 code unit"
    )
    for forbidden in ("Application.International", "DecimalSeparator",
                      "UseSystemSeparators"):
        assert forbidden not in module.code, f"{forbidden} would consult the machine"


def test_36_the_record_builders_take_the_most_likely_flag_from_their_caller() -> None:
    """The resolver owns the distribution vocabulary; the encoder does not infer it."""
    module = _kernel()["modCalcFingerprint"]
    for name in ("CalcFpBuildCostRecord", "CalcFpBuildRiskRecord"):
        joined = " ".join(
            line.strip().rstrip("_") for line in _procedure_body(module, name).splitlines()
        )
        assert "includeMostLikely As Boolean" in joined, f"{name} lacks the flag"


def test_37_the_records_are_sorted_by_permanent_id_before_hashing() -> None:
    body = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpBuildVersionedFingerprint")
    assert body.count("CalcFpSortedRecords(") == 2, (
        "both the cost and the risk section must be ordered"
    )
    sort = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpSortedRecords")
    assert "vbBinaryCompare" in sort


def test_38_the_field_encoding_is_length_prefixed() -> None:
    body = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpField")
    assert "CalcFpUtf16Length(value)" in body
    assert "FP_FIELD_SEPARATOR" in body


# ===========================================================================
# 8. responsibility boundaries
# ===========================================================================
# What each module MAY own, and what it may NOT. The size limits in
# test_phase4_stage_b_source.py are the other half of this pair: size alone
# cannot tell a coherent module from a collapsed one, and these lists can.
MAY_NOT_OWN = {
    "modCalcFactors": {
        # Distribution shapes, headline measures and identities belong one layer up.
        "TriangularMean", "PertMean", "UniformMean", "ExpectedRisk", "Reconcile",
        "AccumulateTotals", "BuildAnnualSeries", "BuildDriverAudit",
        # And the fingerprint belongs one layer sideways.
        "CalcFpDigestStream", "CalcFpReduceDouble", "CalcFpBuildFingerprint",
    },
    "modCalcAnalytical": {
        # The primitives and the exact kernel belong one layer down.
        "SafeAdd", "SafeSubtract", "SafeMultiply", "SafeDivide", "SafeSignedSum",
        "SafeProduct", "BuildKnom", "BuildKpv", "BuildInflationFactors",
        "BuildDiscountFactors",
        # The fingerprint is not an analytical concern.
        "CalcFpDigestStream", "CalcFpReduceDouble", "CalcFpBuildFingerprint",
    },
    "modCalcFingerprint": {
        # No analytical quantity may be computed while encoding one.
        "TriangularMean", "PertMean", "UniformMean", "ExpectedRisk",
        "AccumulateTotals", "BuildAnnualSeries", "Reconcile", "BuildKnom", "BuildKpv",
    },
}

MAY_OWN = {
    "modCalcFactors": ("SafeAdd", "SafeSignedSum", "ExactSumOfProducts",
                       "BuildInflationFactors", "BuildKnom", "IdentityAllowance"),
    "modCalcAnalytical": ("TriangularMean", "DeterministicCentral", "BuildDriverAudit",
                          "AccumulateTotals", "BuildAnnualSeries", "Reconcile"),
    "modCalcFingerprint": ("CalcFpUtf16Length", "CalcFpCanonicalNumber",
                           "CalcFpReduceDouble", "CalcFpDigestStream",
                           "CalcFpBuildFingerprint"),
}


def test_39_each_module_declares_what_it_may_own() -> None:
    for name, expected in MAY_OWN.items():
        declared = set(_kernel()[name].procedures)
        missing = [p for p in expected if p not in declared]
        assert not missing, f"{name} does not own {missing}"


def test_40_no_module_declares_what_another_module_owns() -> None:
    for name, forbidden in MAY_NOT_OWN.items():
        declared = set(_kernel()[name].procedures)
        trespass = sorted(declared & forbidden)
        assert not trespass, f"{name} declares {trespass}, which belongs elsewhere"


def test_41_the_fingerprint_module_computes_no_analytical_quantity() -> None:
    code = _kernel()["modCalcFingerprint"].code
    for token in ("Knom", "Kpv", "MeanValue", "AnnualRow", "Reconcile", "Allowance"):
        assert token not in code, f"the fingerprint module references {token}"


def test_42_the_analytical_module_encodes_nothing() -> None:
    code = _kernel()["modCalcAnalytical"].code
    for token in ("AscW", "CalcFp", "FP_BASE", "FP_MOD_1", "Hex"):
        assert token not in code, f"the analytical module performs encoding: {token}"


# ===========================================================================
# 9. the public API surface, exactly
# ===========================================================================
def test_43_each_module_exposes_exactly_its_whitelisted_public_surface() -> None:
    for name, expected in PUBLIC_SURFACE.items():
        actual = set(_kernel()[name].public_procedures)
        assert actual == expected, (
            f"{name} public surface drifted.\n"
            f"  unexpected: {sorted(actual - expected)}\n"
            f"  missing   : {sorted(expected - actual)}"
        )


def test_44_the_required_minimum_surface_is_inside_the_whitelist() -> None:
    """The whitelist is exact, so it must still contain everything Step 4 requires."""
    required = {
        "IsUsableDouble", "SafeAdd", "SafeSubtract", "SafeMultiply", "SafeDivide",
        "SafeAccumulate", "SafeSignedSum", "SafeProduct", "ExactSumOfProducts",
        "BuildInflationFactors", "BuildDiscountFactors", "BuildKnom", "BuildKpv",
        "ConditioningScaledMagnitude", "IdentityAllowance",
    }
    assert required <= FACTORS_PUBLIC
    assert {"TriangularMean", "PertMean", "UniformMean", "DeterministicCentral",
            "ExpectedRisk"} <= ANALYTICAL_PUBLIC
    for name in ("CalcFpUtf16Length", "CalcFpNormaliseCodeUnit", "CalcFpCanonicalText",
                 "CalcFpCanonicalNumber", "CalcFpCanonicalInteger", "CalcFpReduceDouble",
                 "CalcFpDigestStream", "CalcFpBuildCostRecord", "CalcFpBuildRiskRecord",
                 "CalcFpBuildFingerprint"):
        assert name in FINGERPRINT_PUBLIC


def test_45_every_helper_outside_the_whitelist_is_private() -> None:
    for name, expected in PUBLIC_SURFACE.items():
        module = _kernel()[name]
        helpers = set(module.procedures) - expected
        public = set(module.public_procedures)
        assert not (helpers & public), (
            f"{name} exposes helper(s) {sorted(helpers & public)} that should be Private"
        )


# ===========================================================================
# 10. the locked carry types
# ===========================================================================
def _type_fields(module: VbaModule, name: str) -> list[str]:
    lines = module.raw.splitlines()
    start = next(i for i, line in enumerate(lines) if re.match(rf"^Public Type {name}$", line))
    end = next(i for i in range(start + 1, len(lines)) if lines[i].strip() == "End Type")
    return [
        line.strip().split()[0]
        for line in lines[start + 1 : end]
        if line.strip() and not line.strip().startswith("'")
    ]


def test_46_the_phase_6_carry_types_are_the_locked_ones() -> None:
    factors = _kernel()["modCalcFactors"]
    assert _type_fields(factors, "DriverFactors") == [
        "PermanentId", "IsRisk", "Knom", "Kpv", "Quantity", "Probability",
        "DistKind", "CentralBasis", "MinValue", "MostLikely", "MaxValue",
        "Central", "MeanValue",
    ]
    assert _type_fields(factors, "YearFactors") == [
        "ProjectIndex", "CalendarYear", "DiscountF",
    ]


# ===========================================================================
# 11. TWELVE NEGATIVE CONTROLS
#
# Each plants one defect and asserts the sweep that exists to catch it does.
# A sweep that has silently stopped working would pass every test above and fail
# every test below.
# ===========================================================================
_STUB = 'Attribute VB_Name = "modProbe"\nOption Explicit\n'


def test_nc_01_a_host_call_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function F() As Double\n    F = Application.Sum(1, 2)\nEnd Function\n",
    )
    assert "Application." in boundary_hits(planted)


def test_nc_02_a_worksheet_read_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + 'Public Function F() As Double\n'
        '    F = ThisWorkbook.Worksheets("Calc").Range("A1").Value\nEnd Function\n',
    )
    hits = boundary_hits(planted)
    assert "ThisWorkbook" in hits and "Worksheets" in hits and "Range" in hits


def test_nc_03_a_random_draw_is_caught() -> None:
    planted = _synthetic(
        "modProbe", _STUB + "Public Function F() As Double\n    F = Rnd()\nEnd Function\n"
    )
    assert "Rnd" in boundary_hits(planted)


def test_nc_04_an_excel_object_parameter_type_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function F(ByRef sheet As Worksheet) As Double\n"
        "    F = 0#\nEnd Function\n",
    )
    assert parameter_types(planted) - ALLOWED_PARAMETER_TYPES == {"Worksheet"}


def test_nc_05_on_error_resume_next_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function F() As Double\n    On Error Resume Next\n"
        "    F = 1#\nEnd Function\n",
    )
    assert "On Error Resume Next" in planted.code


def test_nc_06_a_handler_outside_the_four_primitives_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function SafeSignedSum() As Boolean\n"
        "    On Error GoTo ArithmeticFailure\n    SafeSignedSum = True\n"
        "    Exit Function\nArithmeticFailure:\n    SafeSignedSum = False\nEnd Function\n",
    )
    assert handler_owners(planted) - HANDLER_OWNERS == {"SafeSignedSum"}


def test_nc_07_a_pccm_endpoint_is_caught() -> None:
    planted = _synthetic(
        "modProbe", _STUB + "Public Sub PCCM_Calculate()\nEnd Sub\n"
    )
    assert [p for p in planted.procedures if p.startswith("PCCM_")] == ["PCCM_Calculate"]


def test_nc_08_the_native_mod_operator_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function F(ByVal x As Double) As Double\n"
        "    F = x Mod 7\nEnd Function\n",
    )
    assert uses_native_integer_division(planted.code)


def test_nc_09_native_integer_division_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function F(ByVal x As Double) As Double\n"
        "    F = x \\ 7\nEnd Function\n",
    )
    assert uses_native_integer_division(planted.code)


def test_nc_10_a_wide_fixed_point_type_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function F() As Double\n    Dim total As Currency\n"
        "    F = CDec(total)\nEnd Function\n",
    )
    assert wide_type_hits(planted.code) == ["CDec", "Currency"]


def test_nc_11_a_hard_coded_modulus_would_be_caught() -> None:
    """The fingerprint sweep looks for the literal, not for the constant name."""
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function F(ByVal x As Double) As Double\n"
        "    F = x - Fix(x / 2147483647) * 2147483647\nEnd Function\n",
    )
    assert "2147483647" in planted.code


def test_nc_12_public_api_growth_is_caught() -> None:
    """An accidental Public helper is a reviewed entry point that nobody reviewed."""
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function ExactTopBit() As Long\n    ExactTopBit = 0\nEnd Function\n",
    )
    assert set(planted.public_procedures) - FACTORS_PUBLIC == {"ExactTopBit"}


# ===========================================================================
# 13. REVIEW-DISCOVERED DEFECTS
#
# Every test in this section fails against the source as first submitted. Each
# inspects a declaration or an executable body; none searches for a reassuring
# comment.
# ===========================================================================
def _vba_constant(module: VbaModule, name: str) -> float:
    """The value of a `Const <name> As Double = <literal>`, as a Python float."""
    return float(_vba_constant_literal(module, name))


def _vba_constant_literal(module: VbaModule, name: str) -> str:
    """The literal text of a Double constant, with VBA's type suffix removed."""
    for line in module.code_without_string_removal.splitlines():
        match = re.match(
            rf"^\s*(?:Public|Private)\s+Const\s+{name}\s+As\s+Double\s*=\s*(\S+)", line
        )
        if match:
            return match.group(1).rstrip("#!@")
    raise AssertionError(f"{module.name} does not declare {name} As Double")


def record_capacity(module: VbaModule):
    """Evaluate the source's own allocation formula symbolically.

    The `fieldCount = …` and `If includeMostLikely Then fieldCount = …` statements
    are read out of the source and applied, so the test measures what the code
    allocates rather than what a comment claims it allocates.
    """
    body = _procedure_body(module, "CalcFpBuildDriverRecord")
    base = re.search(r"^\s*fieldCount = (\d+) \+ inflationCount \+ weightCount\s*$",
                     body, re.MULTILINE)
    assert base, "the base capacity is not a readable formula"
    bump = re.search(
        r"^\s*If includeMostLikely Then fieldCount = fieldCount \+ (\d+)\s*$",
        body, re.MULTILINE,
    )
    redim = re.search(r"^\s*ReDim fields\(0 To fieldCount - 1\)\s*$", body, re.MULTILINE)
    assert redim, "the allocation must be sized from the computed field count"

    def capacity(include_ml: bool, inflation_count: int, weight_count: int) -> int:
        total = int(base.group(1)) + inflation_count + weight_count
        if include_ml:
            total += int(bump.group(1)) if bump else 0
        return total

    return capacity


def number_field_arguments(module: VbaModule, procedure: str) -> list[str]:
    """The first argument of every `CalcFpNumberField` call, in source order."""
    body = _procedure_body(module, procedure)
    return [m.group(1).strip() for m in re.finditer(r"CalcFpNumberField\(([^,]+),", body)]


def text_field_arguments(module: VbaModule, procedure: str) -> list[str]:
    body = _procedure_body(module, procedure)
    return [m.group(1).strip() for m in re.finditer(r"CalcFpCanonicalText\(([^,)]+)\)", body)]


def variant_declarations(module: VbaModule) -> list[str]:
    """Every `As Variant` in executable code, by line number."""
    return [
        f"{module.name}:{number}"
        for number, statement in logical_statements(module.code)
        if re.search(r"\bAs\s+Variant\b", statement)
    ]


def cleared_fields(module: VbaModule, procedure: str) -> set[str]:
    """The magnitude fields a clear routine assigns by name."""
    body = _procedure_body(module, procedure)
    return {
        match.group(1)
        for match in re.finditer(r"^\s*magnitudes\.(\w+)\s*=\s*0#\s*$", body, re.MULTILINE)
    }


def swallowed_failure_statements(module: VbaModule, procedure: str) -> list[str]:
    """Statements that turn a failed exact evaluation straight into a zero.

    `If Not Exact…(…) Then <var> = 0` treats EVERY exact failure as an accepted
    underflow, including an overflow. That understates a conditioning scale
    without saying so.
    """
    statements = [text for _, text in logical_statements(_procedure_body(module, procedure))]
    hits: list[str] = []
    for index, statement in enumerate(statements):
        if not re.match(r"^If Not \w*Exact\w*\(.*\) Then", statement):
            continue
        tail = statement.split("Then", 1)[1].strip()
        follow = tail or (statements[index + 1] if index + 1 < len(statements) else "")
        if re.match(r"^\w+ = 0#?$", follow):
            hits.append(statement)
    return hits


# --- 13.1 the fingerprint driver-record schema -----------------------------
def test_48_the_shared_record_builder_emits_the_locked_field_order() -> None:
    """ID, Distribution, ONE kind-specific scalar, Min, Max, [ML], FX, inflation
    factors, then profile weights."""
    module = _kernel()["modCalcFingerprint"]
    assert text_field_arguments(module, "CalcFpBuildDriverRecord") == [
        "permanentId", "distribution",
    ]
    assert number_field_arguments(module, "CalcFpBuildDriverRecord") == [
        "kindScalar",
        "minValue",
        "maxValue",
        "mostLikely",
        "fxToSar",
        "inflationFactors(LBound(inflationFactors) + index)",
        "weights(LBound(weights) + index)",
    ]


def test_49_a_cost_record_carries_quantity_and_no_probability() -> None:
    """The opposite kind's multiplicative identity is NOT fingerprinted.

    `Quantity = 1 for risks, Probability = 1 for cost lines` is the in-memory
    DriverFactors carry convention that the calculation and the simulation share.
    It is not the record schema, and writing an identity into the stream would put
    a field in the record that the locked grammar does not have.
    """
    module = _kernel()["modCalcFingerprint"]
    signature = _signature(module, "CalcFpBuildCostRecord")
    assert "quantity As Double" in signature
    assert "probability" not in signature.lower()
    body = _procedure_body(module, "CalcFpBuildCostRecord")
    assert "quantity" in body
    assert not re.search(r"\bprobability\b", body, re.IGNORECASE)
    assert not re.search(r",\s*1#\s*,", body), "a cost record must not encode an identity field"


def test_50_a_risk_record_carries_probability_and_no_quantity() -> None:
    module = _kernel()["modCalcFingerprint"]
    signature = _signature(module, "CalcFpBuildRiskRecord")
    assert "probability As Double" in signature
    assert "quantity" not in signature.lower()
    body = _procedure_body(module, "CalcFpBuildRiskRecord")
    assert "probability" in body
    assert not re.search(r"\bquantity\b", body, re.IGNORECASE)
    assert not re.search(r",\s*1#\s*,", body), "a risk record must not encode an identity field"


def test_51_both_record_builders_take_the_resolved_inflation_factor_vector() -> None:
    """Without it, a change in a referenced inflation factor leaves the record
    unchanged - and a stale result presents itself as current."""
    module = _kernel()["modCalcFingerprint"]
    for name in ("CalcFpBuildCostRecord", "CalcFpBuildRiskRecord",
                 "CalcFpBuildDriverRecord"):
        signature = _signature(module, name)
        assert "inflationFactors() As Double" in signature, f"{name} lacks the vector"
        assert "weights() As Double" in signature, f"{name} lacks the weight vector"


def test_52_inflation_factors_are_encoded_before_the_profile_weights() -> None:
    """Two vectors of the same length in the same record: order is the schema."""
    arguments = number_field_arguments(_kernel()["modCalcFingerprint"],
                                       "CalcFpBuildDriverRecord")
    inflation = next(i for i, a in enumerate(arguments) if a.startswith("inflationFactors"))
    weight = next(i for i, a in enumerate(arguments) if a.startswith("weights"))
    assert inflation < weight, "the inflation vector must precede the weight vector"


def test_53_the_record_capacity_accounts_for_both_vectors_and_the_optional_ml() -> None:
    """The allocation is SIX fixed fields plus both vectors, plus one for ML.

    Permanent ID, Distribution, the kind-specific scalar, Min, Max and FX are the
    six; Most Likely is a seventh only when it is present. A capacity constant
    that folded the optional field in is how a record emitting nine fields came to
    be given eight slots - and the static schema test stayed green while the
    locked one-year Triangular record could not be built at all.
    """
    body = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpBuildDriverRecord")
    assert "inflationCount = UBound(inflationFactors) - LBound(inflationFactors) + 1" in body
    assert "weightCount = UBound(weights) - LBound(weights) + 1" in body
    capacity = record_capacity(_kernel()["modCalcFingerprint"])
    assert capacity(False, 1, 1) == 6 + 1 + 1
    assert capacity(True, 1, 1) == 7 + 1 + 1
    assert capacity(True, 5, 5) == 7 + 5 + 5
    assert capacity(False, 5, 5) == 6 + 5 + 5
    assert "If count <> fieldCount Then Exit Function" in body, (
        "the emitted count must be checked against the capacity the schema asked for"
    )
    assert "CalcFpEncodeRecord(fields, count, record)" in body, (
        "the ENCODED field count is the emitted count, never the array size"
    )


def test_53a_the_capacity_covers_the_locked_triangular_one_year_record() -> None:
    """The reproducer, evaluated symbolically from the source formula.

    ID, Distribution, Quantity, Min, Max, ML, FX, inflation[0], weight[0] is nine
    fields. The allocation must provide nine slots.
    """
    capacity = record_capacity(_kernel()["modCalcFingerprint"])
    emitted = len(number_field_arguments(_kernel()["modCalcFingerprint"],
                                         "CalcFpBuildDriverRecord")) + 2
    assert emitted == 9, "the locked one-year record emits nine fields"
    assert capacity(True, 1, 1) == 9


# --- 13.1b the schema positions must be distinguishable --------------------
# These two tests compute with the accepted PYTHON reference implementation, which
# owns the encoding. They make no statement about VBA: they establish properties
# of the locked stream that the VBA schema above is written to reproduce, and
# they are here because the defect they guard was a schema defect.
def _reference_cost_fields(inflation: float, weight: float) -> tuple:
    """Golden case 1's cost record under the LOCKED schema."""
    from pccm_builder import calc_fingerprint as fp

    return (
        fp.text_field("Triangular"),
        fp.number_field(10), fp.number_field(80), fp.number_field(150),
        fp.number_field(100),
        fp.number_field(1),                 # resolved FX
        fp.number_field(inflation),         # resolved inflation factor
        fp.number_field(weight),            # profiling weight
    )


def test_53b_the_one_year_reference_stream_is_unchanged_by_the_correct_schema() -> None:
    """366 UTF-16 code units and the locked digest, still.

    Golden case 1 has FX = 1, inflation factor = 1 and weight = 1, so its three
    trailing numeric fields are `1, 1, 1` under the locked schema - which is
    exactly why substituting a Probability of 1 for the missing inflation factor
    could hide inside it. Restoring the correct field restores the same stream.
    """
    from pccm_builder import calc_fingerprint as fp
    from pccm_builder.calc_loader import load_calc_contract

    version = load_calc_contract(PCCM_ROOT / "spec" / "calc_contract.yaml").fingerprint_version
    stream = fp.build_canonical_stream(
        version=version,
        header_fields=[fp.number_field(2026), fp.number_field(2026),
                       fp.number_field(1), fp.number_field(0.10)],
        cost_records=[fp.DriverRecord("CL-001", _reference_cost_fields(1, 1))],
    )
    assert fp.utf16_length(stream) == 366
    assert fp.fingerprint(stream) == "50B6EB0E26857EA7"


def test_53c_a_non_identity_inflation_factor_cannot_be_confused_with_anything() -> None:
    """The masking was only possible because every trailing value was 1.

    With a real inflation factor in the slot, a record that dropped it - or that
    put something else there - encodes to a different digest. This is the property
    the one-year fixture could not demonstrate.
    """
    from pccm_builder import calc_fingerprint as fp

    def digest(fields: tuple) -> str:
        return fp.fingerprint(
            fp.encode_section("X", [(fp.text_field("CL-001"),) + fields])
        )

    correct = digest(_reference_cost_fields(1.05, 1))
    identity_substituted = digest(_reference_cost_fields(1, 1))
    dropped = digest(_reference_cost_fields(1.05, 1)[:6] + _reference_cost_fields(1.05, 1)[7:])
    swapped = digest(
        _reference_cost_fields(1.05, 1)[:6]
        + (_reference_cost_fields(1.05, 1)[7], _reference_cost_fields(1.05, 1)[6])
    )
    assert len({correct, identity_substituted, dropped, swapped}) == 4, (
        "the inflation slot must be distinguishable from an identity, from its "
        "absence, and from the weight beside it"
    )


# --- 13.2 the fingerprint version authority --------------------------------
def test_54_the_production_fingerprint_builder_reads_fp_version() -> None:
    module = _kernel()["modCalcFingerprint"]
    signature = _signature(module, "CalcFpBuildFingerprint")
    assert "version As Long" not in signature, (
        "a caller must not be able to select the fingerprint algorithm version"
    )
    assert "FP_VERSION" in _procedure_body(module, "CalcFpBuildFingerprint")


def test_55_the_version_injecting_builder_is_private() -> None:
    """Injectable only for a future migration that compares two encodings."""
    module = _kernel()["modCalcFingerprint"]
    assert "CalcFpBuildVersionedFingerprint" in module.procedures
    assert "CalcFpBuildVersionedFingerprint" not in module.public_procedures


# --- 13.3 the exact binary constants ---------------------------------------
def test_56_the_kernel_constants_are_the_exact_ieee_boundaries() -> None:
    """Compared against the accepted Python authority AND against exact integers.

    These are used by decomposition, by MAX_DOUBLE boundary classification and by
    ties-to-even rounding. They are not documentation values.
    """
    from pccm_builder import calc_numeric

    module = _kernel()["modCalcFactors"]
    assert _vba_constant(module, "TWO_52") == calc_numeric._TWO_52
    assert _vba_constant(module, "TWO_52") == float(2 ** 52)
    assert _vba_constant(module, "MAX_SIGNIFICAND") == calc_numeric._MAX_SIGNIFICAND
    assert _vba_constant(module, "MAX_SIGNIFICAND") == float(2 ** 53 - 1)
    assert _vba_constant(module, "MAX_DOUBLE") == calc_numeric.MAX_DOUBLE
    assert _vba_constant(module, "MAX_DOUBLE") == float((2 ** 53 - 1) * 2 ** 971)


def test_57_the_max_double_literal_is_itself_inside_the_double_range() -> None:
    """A floating-point literal whose MATHEMATICAL value exceeds the greatest
    value representable by its type is statically invalid.

    `float()` would hide this by rounding, so the literal is compared as an exact
    decimal against the exact binary maximum.
    """
    from decimal import Decimal

    literal = _vba_constant_literal(_kernel()["modCalcFactors"], "MAX_DOUBLE")
    assert Decimal(literal) <= Decimal((2 ** 53 - 1) * 2 ** 971), (
        f"the literal {literal} is above the largest representable Double"
    )


def test_58_no_unused_approximate_constant_survives() -> None:
    """A drafted constant that nothing uses is a number nothing checks.

    Counted across all three kernel modules, because a `Public Const` is
    deliberately declared where it is defined and used where it is needed - the
    distribution vocabulary lives with the carry type and is read by the
    analytical layer.
    """
    everywhere = "\n".join(module.code for module in _kernel().values())
    for name, module in _kernel().items():
        for constant in module.constants:
            uses = len(re.findall(rf"(?<![\w.]){constant}(?![\w])", everywhere))
            assert uses > 1, f"{name} declares {constant} but nothing uses it"


# --- 13.4 C1 magnitude ownership -------------------------------------------
def test_59_the_two_magnitude_producers_clear_disjoint_halves() -> None:
    """One structure, two producers, and therefore two ownerships.

    A whole-object clear in each would make a complete structure unreachable:
    whichever ran second would erase what the first captured, and Reconcile needs
    both halves at once.
    """
    module = _kernel()["modCalcAnalytical"]
    headline = cleared_fields(module, "ClearHeadlineMagnitudes")
    annual = cleared_fields(module, "ClearAnnualMagnitudes")
    assert headline == {
        "ANom", "APv", "BNom", "BPv", "CNom", "CPv", "DNom", "DPv", "ENom", "EPv"
    }
    assert annual == {
        "AnnualBaseNom", "AnnualBasePv", "AnnualRiskNom", "AnnualRiskPv",
        "AnnualTotalNom", "AnnualTotalPv",
    }
    assert not (headline & annual), "the two halves overlap"
    declared = set(_type_fields(module, "ReconciliationMagnitudes"))
    assert headline | annual == declared - {"RelativeCoefficient"}, (
        "every magnitude field must belong to exactly one owner"
    )


def test_60_neither_producer_clears_the_whole_structure() -> None:
    module = _kernel()["modCalcAnalytical"]
    assert "ClearMagnitudes" not in module.procedures, (
        "a whole-object clear cannot exist; the two halves have different owners"
    )
    totals = _procedure_body(module, "AccumulateTotals")
    annual = _procedure_body(module, "BuildAnnualSeries")
    assert "ClearHeadlineMagnitudes magnitudes" in totals
    assert "ClearAnnualMagnitudes" not in totals
    assert "ClearAnnualMagnitudes magnitudes" in annual
    assert "ClearHeadlineMagnitudes" not in annual


def test_61_a_conflicting_coefficient_fails_deterministically() -> None:
    """Old magnitudes are never reinterpreted against a tolerance they were not
    measured for."""
    module = _kernel()["modCalcAnalytical"]
    body = _procedure_body(module, "PrepareMagnitudeCoefficient")
    assert "magnitudes.RelativeCoefficient = 0#" in body, (
        "the untouched state must be recognised"
    )
    assert "PrepareMagnitudeCoefficient = (magnitudes.RelativeCoefficient = coefficient)" in body
    for producer in ("AccumulateTotals", "BuildAnnualSeries"):
        assert "PrepareMagnitudeCoefficient(magnitudes, coefficient)" in _procedure_body(
            module, producer
        ), f"{producer} does not verify the coefficient"


# --- 13.5 the empty driver set ---------------------------------------------
def first_bounds_access(module: VbaModule, procedure: str, array: str) -> int:
    """The index of the first statement that reads a bound of, or subscripts, `array`.

    `len(statements)` if it is never touched. This is the line an empty-model
    branch must come BEFORE: an unallocated dynamic array raises on LBound, so a
    guard placed after one can never run.
    """
    statements = [text for _, text in logical_statements(_procedure_body(module, procedure))]
    pattern = re.compile(rf"(?:[LU]Bound\(\s*{array}\b|(?<![\w.]){array}\s*\()")
    # Statement 0 is the declaration itself: `audits() As DriverAudit` names the
    # array without reading it, and a guard cannot precede its own signature.
    for index, statement in enumerate(statements):
        if index and pattern.search(statement):
            return index
    return len(statements)


def empty_branch_index(module: VbaModule, procedure: str, count: str) -> int:
    """The index of the explicit `If <count> = 0 Then` branch."""
    statements = [text for _, text in logical_statements(_procedure_body(module, procedure))]
    for index, statement in enumerate(statements):
        if re.match(rf"^If {count} = 0 Then$", statement):
            return index
    raise AssertionError(f"{procedure} has no explicit zero-{count} branch")


def test_62_the_logical_count_is_a_parameter_at_every_aggregate_boundary() -> None:
    """VBA cannot represent a zero-element array.

    An allocated array always has `UBound >= LBound`, so a count derived from the
    bounds is never zero; an unallocated dynamic array raises Error 9 on `LBound`
    before any emptiness test could run. Deriving the count from the array makes
    the accepted empty model unreachable however the branch is written, so the
    count is passed in.
    """
    module = _kernel()["modCalcAnalytical"]
    assert "auditCount As Long" in _signature(module, "AccumulateTotals")
    assert "driverCount As Long" in _signature(module, "BuildAnnualSeries")
    assert "driverCount As Long" in _signature(module, "Reconcile")
    for procedure, array in (("AccumulateTotals", "audits"),
                             ("BuildAnnualSeries", "drivers"),
                             ("Reconcile", "drivers")):
        body = _procedure_body(module, procedure)
        assert not re.search(rf"count = UBound\({array}\)", body), (
            f"{procedure} still derives its count from the array bounds"
        )


def test_63_every_empty_branch_precedes_any_access_to_its_array() -> None:
    """An empty model must reach each branch WITHOUT touching the array."""
    module = _kernel()["modCalcAnalytical"]
    for procedure, count, arrays in (
        ("AccumulateTotals", "auditCount", ("audits",)),
        ("BuildAnnualSeries", "driverCount", ("drivers", "fxRate", "weights", "inflation")),
        ("Reconcile", "count", ("drivers", "weights")),
    ):
        branch = empty_branch_index(module, procedure, count)
        for array in arrays:
            touch = first_bounds_access(module, procedure, array)
            assert branch < touch, (
                f"{procedure} touches {array} at statement {touch}, before its "
                f"zero-{count} branch at {branch}"
            )


def test_64a_the_empty_branches_produce_the_accepted_empty_model() -> None:
    """Zero totals, one row per applied year, ten identities and no I5 check."""
    module = _kernel()["modCalcAnalytical"]
    totals = _procedure_body(module, "AccumulateTotals")
    assert totals.index("ClearTotals totals") < totals.index("If auditCount = 0 Then")
    annual = _procedure_body(module, "BuildAnnualSeries")
    assert annual.index("ReDim rows(0 To yearCount - 1)") < annual.index("If driverCount = 0 Then")
    assert "rows(offset).ProjectIndex = years(LBound(years) + offset).ProjectIndex" in annual
    reconcile = _procedure_body(module, "Reconcile")
    assert '"no drivers"' not in _procedure_body_raw(module, "Reconcile")
    assert "ReDim checks(0 To 9 + count)" in reconcile, (
        "the ten non-I5 checks must be produced whatever the driver count"
    )
    statements = [text for _, text in logical_statements(reconcile)]
    guard = statements.index("If count = 0 Then")
    order = next(i for i, t in enumerate(statements) if t.startswith("If Not DriverOrder("))
    assert guard < order, "DriverOrder must not be reached with no drivers"


def test_64b_the_private_order_helpers_take_the_logical_count() -> None:
    """A helper that re-derived the count from the array would reintroduce the
    same unreachable branch one level down."""
    module = _kernel()["modCalcAnalytical"]
    for name in ("AuditOrder", "DriverOrder", "CanonicalOrder"):
        assert "count As Long" in _signature(module, name), f"{name} lacks the count"
    for name, array in (("AuditOrder", "audits"), ("DriverOrder", "drivers")):
        body = _procedure_body(module, name)
        assert not re.search(rf"count = UBound\({array}\)", body)


# --- 13.5b the empty sequence identities -----------------------------------
def test_64c_the_sequence_primitives_take_an_explicit_logical_count() -> None:
    """`safe_product([]) == 1.0` is locked behaviour, and it needs a count.

    `If UBound(factors) < LBound(factors) Then` has the same defect as a derived
    driver count: it is unreachable for an allocated array and raises for an
    unallocated one, so the multiplicative identity could never be returned.
    """
    module = _kernel()["modCalcFactors"]
    assert "factorCount As Long" in _signature(module, "SafeProduct")
    assert "termCount As Long" in _signature(module, "SafeSignedSum")
    for procedure in ("SafeProduct", "SafeSignedSum"):
        body = _procedure_body(module, procedure)
        assert "UBound(factors) < LBound(factors)" not in body
        assert "UBound(terms) < LBound(terms)" not in body


def test_64d_the_empty_product_is_one_and_the_empty_sum_is_zero() -> None:
    """Both identities are settled before any bound is read."""
    module = _kernel()["modCalcFactors"]
    for procedure, count, array, identity in (
        ("SafeProduct", "factorCount", "factors", "result = 1#"),
        ("SafeSignedSum", "termCount", "terms", "result = 0#"),
    ):
        statements = [text for _, text in logical_statements(_procedure_body(module, procedure))]
        guard = statements.index(f"If {count} = 0 Then")
        assert statements[guard + 1] == identity, (
            f"{procedure} does not return its identity for an empty sequence"
        )
        assert statements.index(f"If {count} < 0 Then Exit Function") < guard, (
            f"{procedure} must refuse a negative count"
        )
        assert guard < first_bounds_access(module, procedure, array), (
            f"{procedure} reads a bound of {array} before settling the empty case"
        )


def test_64e_every_sequence_call_site_passes_its_logical_count() -> None:
    """No call site may fall back to the allocated capacity."""
    expected = {
        ("modCalcFactors", "BuildFactor"): {"SafeProduct(group, width,",
                                            "SafeSignedSum(terms, count,"},
        ("modCalcAnalytical", "ExpectedRisk"): {"SafeProduct(group, 3,"},
        ("modCalcAnalytical", "TripleProduct"): {"SafeProduct(group, 3,"},
        ("modCalcAnalytical", "AnnualSeries"): {"SafeSignedSum(terms, count,"},
        ("modCalcAnalytical", "SumMeasure"): {"SafeSignedSum(terms, count,"},
    }
    for (module_name, procedure), fragments in expected.items():
        body = _procedure_body(_kernel()[module_name], procedure)
        for fragment in fragments:
            assert fragment in body, f"{module_name}.{procedure} lacks {fragment}"
    annual = _procedure_body(_kernel()["modCalcAnalytical"], "BuildAnnualSeries")
    assert "SafeProduct(group, ANNUAL_FACTOR_COUNT," in annual
    reconcile = _procedure_body(_kernel()["modCalcAnalytical"], "Reconcile")
    assert "SafeSignedSum(series, yearCount, total)" in reconcile
    assert "SafeSignedSum(series, UBound(series) + 1, total)" in reconcile


# --- 13.5c the exact quotient divisor boundary -----------------------------
def test_64f_an_unsupported_divisor_is_refused_before_the_exact_division() -> None:
    """`ExactQuotientOfSum` is Public and installs no error handler.

    A divisor of zero would otherwise reach a raw division inside it. The
    contract is exactly the three convex-statistic denominators.
    """
    module = _kernel()["modCalcFactors"]
    statements = [
        text for _, text in
        logical_statements(_procedure_body(module, "ExactQuotientOfSum"))
    ]
    guard = statements.index(
        "If divisor <> 2# And divisor <> 3# And divisor <> 6# Then Exit Function"
    )
    division = next(i for i, t in enumerate(statements) if "ExactDivideSmall(" in t)
    assert guard < division, "the divisor must be validated before the division"
    assert guard < next(i for i, t in enumerate(statements) if "ExactSumOf(" in t), (
        "an unsupported divisor should be refused before any work is done"
    )
    assert only_supported_divisors_are_passed(), (
        "a caller passes a divisor outside the documented contract"
    )


def only_supported_divisors_are_passed() -> bool:
    """Every ExactQuotientOfSum call site's divisor is 2, 3 or 6."""
    for module in _kernel().values():
        for _, statement in logical_statements(module.code):
            for match in re.finditer(r"ExactQuotientOfSum\(([^)]*)\)", statement):
                arguments = [a.strip() for a in match.group(1).split(",")]
                if len(arguments) == 4 and arguments[2] not in {"2#", "3#", "6#", "divisor"}:
                    return False
    return True


def test_64g_the_convex_statistics_pass_only_supported_divisors() -> None:
    module = _kernel()["modCalcAnalytical"]
    divisors = {
        "TriangularMean": "3#", "PertMean": "6#", "UniformMean": "2#",
    }
    for procedure, divisor in divisors.items():
        body = _procedure_body(module, procedure)
        assert re.search(rf"ConvexFinish\(.*, {re.escape(divisor)}, result\)", body), (
            f"{procedure} does not pass {divisor}"
        )


# --- 13.5d public surface discipline ---------------------------------------
def test_64h_every_public_helper_has_a_cross_module_caller_or_a_stated_reason() -> None:
    """Scans references, not comments.

    A Public name with no caller outside its own module and no entry in the
    documented exception set is accidental API growth.
    """
    modules = _kernel()
    unexplained: list[str] = []
    for name, module in modules.items():
        others = "\n".join(other.code for label, other in modules.items() if label != name)
        for procedure in module.public_procedures:
            if procedure in PUBLIC_WITHOUT_CROSS_MODULE_CALLER:
                continue
            called = re.search(rf"(?<![\w.]){procedure}\s*\(", others)
            if not called:
                unexplained.append(f"{name}.{procedure}")
    assert not unexplained, (
        "Public with no cross-module caller and no documented reason: "
        f"{sorted(unexplained)}"
    )


def test_64i_the_three_reviewed_helpers_are_private() -> None:
    """None had a cross-module caller, and none was part of the diagnostic surface."""
    analytical = _kernel()["modCalcAnalytical"]
    fingerprint = _kernel()["modCalcFingerprint"]
    for name in ("DistributionMean", "CanonicalOrder"):
        assert name in analytical.procedures, f"{name} must keep its semantics"
        assert name not in analytical.public_procedures, f"{name} must be Private"
    assert "CalcFpNumberField" in fingerprint.procedures
    assert "CalcFpNumberField" not in fingerprint.public_procedures


# --- 13.6 conditioning underflow versus overflow ---------------------------
def test_64_a_failed_exact_evaluation_is_never_turned_into_a_zero() -> None:
    """C1's narrow exception is UNDERFLOW, not "any failure".

    A final conditioning magnitude outside Double range must be reported. Silently
    recording zero understates the scale, and a tolerance may never be narrowed by
    accident any more than it may be widened by one.
    """
    module = _kernel()["modCalcFactors"]
    for procedure in ("ConditioningScaledProduct", "ConditioningScaledMagnitude",
                      "ConditioningScaledExact"):
        assert swallowed_failure_statements(module, procedure) == [], (
            f"{procedure} converts an exact failure into a zero"
        )


def test_65_the_conditioning_path_uses_the_underflow_to_zero_policy() -> None:
    """One kernel, two policies, and the policy is a parameter."""
    module = _kernel()["modCalcFactors"]
    public = _procedure_body(module, "ExactSumOfProducts")
    assert re.search(r"ExactSumOfProductsCore\(.*groupCount, False, result\)", public, re.S), (
        "model arithmetic must refuse an underflow"
    )
    conditioning = _procedure_body(module, "ConditioningScaledExact")
    assert re.search(r"ExactSumOfProductsCore\(.*groupCount, _?\s*True, scaled\)",
                     conditioning, re.S), (
        "conditioning metadata must be allowed to underflow to zero"
    )


def test_66_conditioning_does_not_retry_the_multiplication_that_failed() -> None:
    """An unchecked arithmetic retry after SafeMultiply refused is not a rescue."""
    body = _procedure_body(_kernel()["modCalcFactors"], "ConditioningScaledMagnitude")
    assert not re.search(r"^\s*scaled = coefficient \* magnitude\s*$", body, re.MULTILINE), (
        "the failed multiplication must not be repeated outside a safe primitive"
    )
    assert "ConditioningScaledExact(group, scaled)" in body


def test_67_an_overflow_cannot_be_labelled_as_an_accepted_underflow() -> None:
    """The range refusals in RoundExact do not consult the underflow flag."""
    statements = [
        text for _, text in
        logical_statements(_procedure_body(_kernel()["modCalcFactors"], "RoundExact"))
    ]
    overflow = [s for s in statements if "Exit Function" in s and (
        "exponent > 1023" in s or "MAX_SIGNIFICAND" in s
    )]
    assert len(overflow) == 3, f"expected three range refusals, found {overflow}"
    for statement in overflow:
        assert "underflowToZero" not in statement, (
            "an out-of-range magnitude may never be accepted as an underflow"
        )


# --- 13.7 typed numerical containers ---------------------------------------
def test_68_no_kernel_module_declares_a_variant() -> None:
    """The numerical boundary is Double, Long, Boolean, String and typed records.

    A Variant container was never an approved resolution of this constraint, and
    the allowed-type list is not the place to widen it.
    """
    for name, module in _kernel().items():
        assert variant_declarations(module) == [], (
            f"{name} declares a Variant: {variant_declarations(module)}"
        )
    assert "Variant" not in ALLOWED_PARAMETER_TYPES


def test_69_the_sum_of_products_container_is_a_flat_typed_vector() -> None:
    module = _kernel()["modCalcFactors"]
    signature = _signature(module, "ExactSumOfProducts")
    for part in ("factors() As Double", "groupStarts() As Long",
                 "groupLengths() As Long", "groupCount As Long"):
        assert part in signature, f"ExactSumOfProducts lacks {part}"
    assert "factors() As Double" in _signature(module, "SafeProduct")
    assert "factors() As Double" in _signature(module, "ConditioningScaledProduct")


def test_70_every_exact_rescue_site_builds_the_typed_vector() -> None:
    """Both production callers describe their groups with the same two Long
    vectors, so no site can drift back to an untyped container."""
    for module_name, procedure in (("modCalcFactors", "BuildFactor"),
                                   ("modCalcAnalytical", "AnnualSeries")):
        body = _procedure_body(_kernel()[module_name], procedure)
        assert "flat() As Double" in body or "Dim flat() As Double" in body
        assert "starts() As Long" in body
        assert "lengths() As Long" in body
        assert re.search(r"ExactSumOfProducts\(flat, starts, lengths, count, result\)", body)


def _signature(module: VbaModule, name: str) -> str:
    """One procedure's declaration, with continuations joined."""
    for _, statement in logical_statements(module.code_without_string_removal):
        if re.match(rf"^\s*(Public |Private |Friend )?(Static )?(Sub|Function)\s+{name}\b",
                    statement):
            return re.sub(r"\s+", " ", statement)
    raise AssertionError(f"{module.name} does not declare {name}")


def _procedure_body_raw(module: VbaModule, name: str) -> str:
    """A procedure body with string literals INTACT, for literal-text checks."""
    lines = module.code_without_string_removal.splitlines()
    start = next(
        i for i, line in enumerate(lines)
        if re.match(rf"^\s*(Public |Private )?(Static )?(Sub|Function)\s+{name}\b", line)
    )
    end = next(i for i in range(start + 1, len(lines))
               if re.match(r"^End (Sub|Function)", lines[i]))
    return "\n".join(lines[start:end])


# --- 13.8 function results -------------------------------------------------
_DECLARATION = re.compile(r"^(Public |Private |Friend )?(Static )?(Sub|Function)\s+(\w+)")
_ASSIGNMENT = re.compile(r"^([A-Za-z_]\w*)\s*=\s*[^=]")
_NOT_AN_ASSIGNMENT = re.compile(
    r"^(If|ElseIf|Dim|Const|ReDim|For|Do|While|Set|Select|Case|Next|Loop|End|Exit)\b",
    re.IGNORECASE,
)


def function_result_assignments(module: VbaModule) -> dict[str, set[str]]:
    """For each Function in `module`, the plain identifiers its body assigns.

    A VBA function returns by assigning to its own name, so an assignment to the
    WRONG name compiles cleanly, returns the default, and is invisible to every
    test that only reads signatures. Array-element and UDT-member assignments are
    excluded: `fields(0) = …` and `check.Label = …` are not return statements.
    """
    functions: dict[str, set[str]] = {}
    current: str | None = None
    for _, statement in logical_statements(module.code):
        declaration = _DECLARATION.match(statement)
        if declaration:
            current = declaration.group(4) if declaration.group(3).lower() == "function" else None
            if current:
                functions[current] = set()
            continue
        if current is None:
            continue
        if not _NOT_AN_ASSIGNMENT.match(statement):
            assignment = _ASSIGNMENT.match(statement)
            if assignment:
                functions[current].add(assignment.group(1))
        # `If <cond> Then <name> = …` is a return on a single line.
        for tail in re.finditer(r"\bThen\s+([A-Za-z_]\w*)\s*=\s*[^=]", statement):
            functions[current].add(tail.group(1))
    return functions


def foreign_result_assignments(module: VbaModule) -> list[str]:
    """Functions that assign the name of a DIFFERENT function in the same module."""
    assignments = function_result_assignments(module)
    names = set(assignments)
    return sorted(
        f"{name} -> {sorted((assigned & names) - {name})}"
        for name, assigned in assignments.items()
        if (assigned & names) - {name}
    )


def functions_never_assigning_their_own_result(module: VbaModule) -> list[str]:
    return sorted(
        name for name, assigned in function_result_assignments(module).items()
        if name not in assigned
    )


def test_71_every_function_returns_through_its_own_result_name() -> None:
    """A mechanical gate, so this class of typo is not rediscovered one at a time.

    VBA has no compiler here to notice that a procedure assigns a name that is not
    its own: the assignment is legal, the caller silently receives the default,
    and every signature-level test still passes. This checks the one thing that
    catches it - that each Function names itself somewhere on a successful path.

    It is deliberately NOT a control-flow proof. A failure-only path need not
    assign anything.
    """
    for name, module in _kernel().items():
        missing = functions_never_assigning_their_own_result(module)
        assert not missing, f"{name}: these Functions never assign their own result: {missing}"


def test_72_no_function_assigns_another_functions_result_name() -> None:
    for name, module in _kernel().items():
        foreign = foreign_result_assignments(module)
        assert not foreign, f"{name}: assignment to another Function's result: {foreign}"


def test_73_the_private_fingerprint_builder_returns_its_own_result() -> None:
    """The specific instance the gate above generalises.

    `CalcFpBuildVersionedFingerprint` computes the digest and the public wrapper
    returns whatever it returns, so assigning the WRAPPER's name inside the helper
    discards the result entirely.
    """
    module = _kernel()["modCalcFingerprint"]
    body = _procedure_body(module, "CalcFpBuildVersionedFingerprint")
    assert "CalcFpBuildVersionedFingerprint = CalcFpDigestStream(stream, result)" in body
    assert "CalcFpBuildFingerprint =" not in body, (
        "the helper must not assign the public wrapper's result name"
    )
    wrapper = _procedure_body(module, "CalcFpBuildFingerprint")
    assert "CalcFpBuildFingerprint = CalcFpBuildVersionedFingerprint(" in wrapper


# ===========================================================================
# 14. NEGATIVE CONTROLS FOR THE REVIEW-DISCOVERED DEFECTS
# ===========================================================================
def test_nc_13_reintroducing_probability_into_a_cost_record_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CalcFpBuildDriverRecord() As Boolean\n"
        "    If Not CalcFpNumberField(quantity, decimalSeparator, fields(count)) Then Exit Function\n"
        "    If Not CalcFpNumberField(minValue, decimalSeparator, fields(count)) Then Exit Function\n"
        "    If Not CalcFpNumberField(maxValue, decimalSeparator, fields(count)) Then Exit Function\n"
        "    If Not CalcFpNumberField(probability, decimalSeparator, fields(count)) Then Exit Function\n"
        "    If Not CalcFpNumberField(fxToSar, decimalSeparator, fields(count)) Then Exit Function\n"
        "End Function\n",
    )
    assert "probability" in number_field_arguments(planted, "CalcFpBuildDriverRecord")


def test_nc_14_dropping_the_inflation_vector_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CalcFpBuildDriverRecord() As Boolean\n"
        "    If Not CalcFpNumberField(fxToSar, decimalSeparator, fields(count)) Then Exit Function\n"
        "    If Not CalcFpNumberField(weights(LBound(weights) + index), decimalSeparator, _\n"
        "                             fields(count)) Then Exit Function\n"
        "End Function\n",
    )
    arguments = number_field_arguments(planted, "CalcFpBuildDriverRecord")
    assert not any(a.startswith("inflationFactors") for a in arguments)


def test_nc_15_swapping_the_inflation_and_weight_loops_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CalcFpBuildDriverRecord() As Boolean\n"
        "    If Not CalcFpNumberField(weights(LBound(weights) + index), decimalSeparator, _\n"
        "                             fields(count)) Then Exit Function\n"
        "    If Not CalcFpNumberField(inflationFactors(LBound(inflationFactors) + index), _\n"
        "                             decimalSeparator, fields(count)) Then Exit Function\n"
        "End Function\n",
    )
    arguments = number_field_arguments(planted, "CalcFpBuildDriverRecord")
    inflation = next(i for i, a in enumerate(arguments) if a.startswith("inflationFactors"))
    weight = next(i for i, a in enumerate(arguments) if a.startswith("weights"))
    assert weight < inflation, "the planted swap must be visible to the ordering check"


def test_nc_16_a_caller_selected_fingerprint_version_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function CalcFpBuildFingerprint(ByVal version As Long, _\n"
        "                                       ByRef result As String) As Boolean\n"
        "End Function\n",
    )
    assert "version As Long" in _signature(planted, "CalcFpBuildFingerprint")


def test_nc_17_a_rounded_binary_constant_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Const TWO_52 As Double = 4503599627370500#\n"
        "Private Const MAX_SIGNIFICAND As Double = 9007199254740990#\n",
    )
    assert _vba_constant(planted, "TWO_52") != float(2 ** 52)
    assert _vba_constant(planted, "MAX_SIGNIFICAND") != float(2 ** 53 - 1)


def test_nc_18_an_out_of_range_max_double_literal_is_caught() -> None:
    from decimal import Decimal

    planted = _synthetic(
        "modProbe", _STUB + "Public Const MAX_DOUBLE As Double = 1.79769313486232E+308\n"
    )
    literal = _vba_constant_literal(planted, "MAX_DOUBLE")
    assert Decimal(literal) > Decimal((2 ** 53 - 1) * 2 ** 971)


def test_nc_19_a_whole_object_magnitude_clear_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub ClearMagnitudes(ByRef magnitudes As ReconciliationMagnitudes)\n"
        "    Dim blank As ReconciliationMagnitudes\n    magnitudes = blank\nEnd Sub\n",
    )
    assert "ClearMagnitudes" in planted.procedures
    assert cleared_fields(planted, "ClearMagnitudes") == set(), (
        "a whole-object clear names no field, so it can own no half"
    )


def test_nc_20_an_explicit_no_driver_refusal_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function Reconcile() As Boolean\n    If count < 1 Then\n"
        '        detail = "no drivers"\n        Exit Function\n    End If\nEnd Function\n',
    )
    assert '"no drivers"' in _procedure_body_raw(planted, "Reconcile")


def test_nc_21_an_exact_failure_turned_into_zero_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function F() As Boolean\n"
        "    If Not ExactSumOfProducts(groups, scaled) Then\n        scaled = 0#\n"
        "    End If\nEnd Function\n",
    )
    assert swallowed_failure_statements(planted, "F") != []


def test_nc_31_a_foreign_function_result_assignment_is_caught() -> None:
    """The exact defect: the helper assigns the wrapper's name."""
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function CalcFpBuildFingerprint() As Boolean\n"
        "    CalcFpBuildFingerprint = CalcFpBuildVersionedFingerprint(stream, result)\n"
        "End Function\n"
        "Private Function CalcFpBuildVersionedFingerprint() As Boolean\n"
        "    CalcFpBuildFingerprint = CalcFpDigestStream(stream, result)\n"
        "End Function\n",
    )
    assert functions_never_assigning_their_own_result(planted) == [
        "CalcFpBuildVersionedFingerprint"
    ]
    assert foreign_result_assignments(planted) == [
        "CalcFpBuildVersionedFingerprint -> ['CalcFpBuildFingerprint']"
    ]


def test_nc_32_the_result_gate_does_not_fire_on_a_correct_function() -> None:
    """The other direction: a function that does name itself must pass.

    Array-element and member assignments are not return statements and must not
    be mistaken for one.
    """
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function F(ByRef fields() As String) As Boolean\n"
        "    Dim check As IdentityCheck\n"
        "    fields(0) = \"x\"\n"
        "    check.Label = \"y\"\n"
        "    If ok Then F = True\n"
        "End Function\n",
    )
    assert functions_never_assigning_their_own_result(planted) == []
    assert foreign_result_assignments(planted) == []


def test_nc_33_a_host_marker_equality_gate_is_caught() -> None:
    """Restoring the gate must be visible to the sweep."""
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function CalcFpCanonicalNumber() As Boolean\n"
        "    marker = CalcFpMarkerIndex(text)\n"
        "    If marker = 0 Then Exit Function\n"
        "    If Mid$(text, marker, 1) <> decimalSeparator Then Exit Function\n"
        "    CalcFpCanonicalNumber = True\n"
        "End Function\n",
    )
    assert host_marker_equality_gate(planted, "CalcFpCanonicalNumber") == [
        "If Mid$(text, marker, 1) <> decimalSeparator Then Exit Function"
    ]


def test_nc_23_a_capacity_formula_that_forgets_most_likely_is_caught() -> None:
    """The submitted formula, evaluated for the locked one-year Triangular shape."""
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CalcFpBuildDriverRecord() As Boolean\n"
        "    fieldCount = 5 + inflationCount + weightCount\n"
        "    ReDim fields(0 To fieldCount - 1)\n"
        "End Function\n",
    )
    capacity = record_capacity(planted)
    assert capacity(True, 1, 1) == 7, "the planted formula allocates seven slots"
    assert capacity(True, 1, 1) < 9, "and the locked record needs nine"


def test_nc_24_a_missing_emitted_count_guard_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CalcFpBuildDriverRecord() As Boolean\n"
        "    CalcFpBuildDriverRecord = CalcFpEncodeRecord(fields, count, record)\n"
        "End Function\n",
    )
    body = _procedure_body(planted, "CalcFpBuildDriverRecord")
    assert "If count <> fieldCount Then Exit Function" not in body


def test_nc_25_a_count_derived_from_array_bounds_is_caught() -> None:
    """The pattern that makes an empty model unreachable, in all three producers."""
    for procedure, array, count in (("AccumulateTotals", "audits", "auditCount"),
                                    ("BuildAnnualSeries", "drivers", "driverCount"),
                                    ("Reconcile", "drivers", "driverCount")):
        planted = _synthetic(
            "modProbe",
            _STUB + f"Public Function {procedure}(ByRef {array}() As DriverAudit) As Boolean\n"
            f"    count = UBound({array}) - LBound({array}) + 1\n"
            "    If count < 1 Then\n        Exit Function\n    End If\n"
            "End Function\n",
        )
        body = _procedure_body(planted, procedure)
        assert re.search(rf"count = UBound\({array}\)", body), (
            "the derived-count pattern must be visible to the sweep"
        )
        assert count not in _signature(planted, procedure), (
            "and the planted version has no logical count parameter"
        )


def test_nc_26_an_empty_branch_placed_after_a_bounds_read_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function AccumulateTotals(ByRef audits() As DriverAudit, _\n"
        "                                 ByVal auditCount As Long) As Boolean\n"
        "    count = UBound(audits) - LBound(audits) + 1\n"
        "    If auditCount = 0 Then\n        AccumulateTotals = True\n"
        "        Exit Function\n    End If\n"
        "End Function\n",
    )
    branch = empty_branch_index(planted, "AccumulateTotals", "auditCount")
    touch = first_bounds_access(planted, "AccumulateTotals", "audits")
    assert touch < branch, "the planted order must be visible to the sweep"


def test_nc_27_an_unreachable_empty_product_branch_is_caught() -> None:
    """`If UBound(factors) < LBound(factors)` can never be true for an allocated
    array, and raises for an unallocated one."""
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function SafeProduct(ByRef factors() As Double) As Boolean\n"
        "    If UBound(factors) < LBound(factors) Then\n        result = 1#\n"
        "        SafeProduct = True\n        Exit Function\n    End If\n"
        "End Function\n",
    )
    body = _procedure_body(planted, "SafeProduct")
    assert "UBound(factors) < LBound(factors)" in body
    assert "factorCount As Long" not in _signature(planted, "SafeProduct")


def test_nc_28_a_count_less_signed_sum_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function SafeSignedSum(ByRef terms() As Double, _\n"
        "                              ByRef result As Double) As Boolean\n"
        "End Function\n",
    )
    assert "termCount As Long" not in _signature(planted, "SafeSignedSum")


def test_nc_29_an_unguarded_exact_divisor_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function ExactQuotientOfSum(ByVal divisor As Double) As Boolean\n"
        "    If Not ExactSumOf(terms, termCount, exact) Then Exit Function\n"
        "    remainder = ExactDivideSmall(guarded, divisor, quotient)\n"
        "End Function\n",
    )
    statements = [
        text for _, text in
        logical_statements(_procedure_body(planted, "ExactQuotientOfSum"))
    ]
    assert not any("divisor <> 2#" in t for t in statements), (
        "the planted version reaches the division with no divisor guard"
    )


def test_nc_30_an_accidentally_public_helper_is_caught() -> None:
    """No cross-module caller and no documented reason is accidental growth."""
    for name in ("DistributionMean", "CanonicalOrder", "CalcFpNumberField"):
        assert name not in PUBLIC_WITHOUT_CROSS_MODULE_CALLER, (
            f"{name} must not be excused; it has no external consumer"
        )
    planted = _synthetic(
        "modProbe", _STUB + "Public Function DistributionMean() As Boolean\nEnd Function\n"
    )
    assert "DistributionMean" in planted.public_procedures
    others = "\n".join(m.code for m in _kernel().values())
    assert not re.search(r"(?<![\w.])DistributionMean\s*\(", "\n".join(
        m.code for name, m in _kernel().items() if name != "modCalcAnalytical"
    )), "DistributionMean genuinely has no cross-module caller"
    assert others  # the scan operated on real source, not on an empty string


def test_nc_22_a_variant_numerical_container_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function SafeProduct(ByRef factors As Variant) As Boolean\n"
        "    Dim groups() As Variant\nEnd Function\n",
    )
    assert len(variant_declarations(planted)) == 2


# ===========================================================================
# 12. this suite makes no runtime claim
# ===========================================================================
def test_47_no_test_in_this_file_claims_that_vba_ran() -> None:
    """A guard on the suite's own language.

    The one thing a static suite must never do is describe itself as evidence of
    execution. If a future edit asserts that the VBA yielded a value, or that a
    parity result has been established, this test fails before the claim can reach
    a reviewer.

    The forbidden phrases are ASSEMBLED FROM PARTS rather than written out, so
    this guard does not trip over its own wording - and so that adding a phrase
    here can never accidentally make the file contain it.
    """
    text = Path(__file__).read_text(encoding="utf-8")
    banned = (
        ("VBA", "produced"), ("VBA", "computed"), ("VBA", "returned"),
        ("VBA", "evaluated"), ("parity", "is proven"), ("parity", "proven"),
        ("executed", "the VBA"), ("ran", "the VBA"),
    )
    for parts in banned:
        claim = " ".join(parts)
        assert claim not in text, f"this suite must not make that claim: {parts}"
    assert "NO VBA IS EXECUTED HERE" in text


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
