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
ALLOWED_PARAMETER_TYPES = {
    "Double", "Long", "Boolean", "String", "Variant",
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
    "TriangularMean", "PertMean", "UniformMean", "DistributionMean",
    "DeterministicCentral", "ExpectedRisk",
    "CanonicalOrder", "BuildDriverAudit", "AccumulateTotals",
    "BuildAnnualSeries", "Reconcile", "AllIdentitiesHold",
}

FINGERPRINT_PUBLIC = {
    "CalcFpUtf16Length", "CalcFpNormaliseCodeUnit", "CalcFpCanonicalText",
    "CalcFpCanonicalNumber", "CalcFpCanonicalInteger", "CalcFpNumberField",
    "CalcFpReduceDouble", "CalcFpDigestStream", "CalcFpBuildCostRecord",
    "CalcFpBuildRiskRecord", "CalcFpBuildFingerprint",
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
    """The accepted module split is deliberate.

    A modCalcMath, modCalcTypes or modCalcExact would be a fourth production
    module that no review accepted, and the split would stop meaning anything.
    """
    on_disk = set(_modules())
    assert on_disk == set(PHASE4_MODULES) | set(KERNEL_MODULES), (
        f"unexpected hand-written module inventory: {sorted(on_disk)}"
    )


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
    everything = "\n".join(m.raw for m in _modules().values())
    for deferred in (
        "modCalcResolve", "modCalcCheck", "modCalcReport",
        "PCCM_Calculate", "PCCM_CalculationStatus", "PCCM_CalculationAttemptResult",
        "PCCM_CalculationAttemptDetail", "PCCM_CalculationFingerprint",
        "PCCM_CurrentInputFingerprint",
    ):
        assert deferred not in everything, f"{deferred} belongs to a later step"


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
    assert "CalcFpMarkerIndex(text)" in body


def test_36_the_record_builders_take_the_most_likely_flag_from_their_caller() -> None:
    """The resolver owns the distribution vocabulary; the encoder does not infer it."""
    module = _kernel()["modCalcFingerprint"]
    for name in ("CalcFpBuildCostRecord", "CalcFpBuildRiskRecord"):
        joined = " ".join(
            line.strip().rstrip("_") for line in _procedure_body(module, name).splitlines()
        )
        assert "includeMostLikely As Boolean" in joined, f"{name} lacks the flag"


def test_37_the_records_are_sorted_by_permanent_id_before_hashing() -> None:
    body = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpBuildFingerprint")
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
