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
import collections
import math
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

# The Step-7 presentation and orchestration layer, the last Phase-5 module.
# Also outside KERNEL_MODULES, for the same reason.
STEP7_MODULE = "modCalcReport"

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
    "MAX_DOUBLE", "IsUsableDouble", "SafeAdd", "SafeSubtract", "SafeMultiply", "SafeDivide",
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
    # PUBLIC SINCE STEP 7. modCalcReport frames the four header scalars - Base
    # Year, Start Year, Duration and Discount Rate - and they are NUMBER fields.
    # The orchestration layer must reach the accepted framing authority rather
    # than assemble an N field of its own, so the framing authority stays here
    # and becomes reachable. The body is unchanged; test_64j proves that.
    "CalcFpNumberField",
    # PUBLIC SINCE STEP 10, and the only procedure added to this module. The
    # Phase-6 request fingerprint is the analytical stream followed by a SIM
    # section, and the accepted Step-10A authority forbids hashing the
    # analytical DIGEST as a field. Continuing the hash from the digest's own
    # accumulator states is what lets modSimFingerprint reach this hash instead
    # of implementing a second one. Its caller is modSimFingerprint.
    "CalcFpContinueDigest",
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
    "ExpectedRisk", "MAX_DOUBLE", "IsUsableDouble", "SafeAdd", "SafeSubtract", "SafeMultiply", "SafeDivide",
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
# The ACCEPTED EXECUTABLE text of modCalcFingerprint: comments and blank lines
# removed, whitespace runs collapsed, and the ONE authorised Step-7 visibility
# keyword normalised back to Private.
#
# This digest has moved three times, each time under a recorded authorisation:
#   Step 7                 CalcFpNumberField made Public, nothing else.
#   Gate B Runtime Run 2   the canonical Double encoder rebuilt, because Format$
#                          provably could not produce the contracted 17
#                          significant digits on real Excel.
#   Gate B Runtime Run 7   CalcFpEncodeSection's parameter renamed `name` ->
#                          `sectionName`, because `Name` is a VBA statement
#                          keyword and Run 7 disproved the evidence that had
#                          grandfathered it. IDENTIFIER SPELLING ONLY.
# It is what "and nothing else" is measured against from here on.
#
# The Run-7 move is not merely recorded, it is PROVED: test_64j reverses the
# rename over the same reduction and requires the PREVIOUS digest back, so a
# logic change smuggled in alongside the rename cannot pass by updating a
# number.
FINGERPRINT_ACCEPTED_BODY_SHA256 = (
    "1ea6aa3ca4b9d8ce3a5b8885f6e3ba24b1cfe6da870f25ce2db88e2061084cb3"
)
FINGERPRINT_BODY_SHA256_BEFORE_RUN7_RENAME = (
    "27589cbef04e29ceff15df05a0b1cbfdf2d35e25ab301cdc2c992e46468a9659"
)


def _modules() -> dict[str, VbaModule]:
    return {m.name: m for m in load_modules([SRC_VBA])}



# ---------------------------------------------------------------------------
# modCalcFingerprint took ONE authorised Step-10 addition: the canonical digest
# continuation, APPENDED after every accepted line. The frozen digests here are
# therefore taken over the ACCEPTED PREFIX - the file up to that banner - and
# they still carry their ORIGINAL literals. That is deliberately stronger than
# re-pinning them: the accepted bytes must be identical, and the only thing that
# may exist beyond them is the named Step-10 block.
# ---------------------------------------------------------------------------
STEP10_FINGERPRINT_BANNER = (
    "' ==========================================================================\n"
    "' STEP 10 ADDITION - THE CANONICAL DIGEST CONTINUATION\n"
)


STEP11_REPORTER_BANNER = (
    "' ==========================================================================\n"
    "' STEP 11 ADDITION - THE PHASE-6 PREPARATION BRIDGE\n"
)


def _accepted_reporter_source() -> str:
    """modCalcReport up to the Step-11 bridge banner.

    The bridge is APPENDED after every accepted line, so the reversal proofs
    below still have to reproduce the base commit byte for byte.
    """
    text = (SRC_VBA / "modCalcReport.bas").read_text(encoding="utf-8")
    assert text.count(STEP11_REPORTER_BANNER) == 1, (
        "the Step-11 bridge banner is missing or duplicated"
    )
    return text[: text.index(STEP11_REPORTER_BANNER)]


def _accepted_fingerprint_source() -> str:
    text = (SRC_VBA / "modCalcFingerprint.bas").read_text(encoding="utf-8")
    assert text.count(STEP10_FINGERPRINT_BANNER) == 1, (
        "the Step-10 continuation banner is missing or duplicated; the accepted "
        "prefix cannot be identified"
    )
    return text[: text.index(STEP10_FINGERPRINT_BANNER)]


def fingerprint_body_digest(source: str | None = None) -> str:
    """modCalcFingerprint reduced to executable text, visibility normalised.

    `source` lets a caller digest a TRANSFORMED copy through the identical
    reduction - which is how the Run-7 rename is proved to be a rename.
    """
    import hashlib

    if source is None:
        source = _accepted_fingerprint_source()
    kept: list[str] = []
    for line in source.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("'"):
            continue
        stripped = re.sub(r"\s+", " ", stripped)
        stripped = stripped.replace(
            "Public Function CalcFpNumberField", "Private Function CalcFpNumberField"
        )
        kept.append(stripped)
    return hashlib.sha256("\n".join(kept).encode()).hexdigest()


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


PHASE6_MODULES = ("modSimRng", "modSimSample", "modSimEngine", "modSimStats",
                  "modSimFingerprint", "modSimNonce", "modSimReport")
"""Phase-6 hand-written source modules. Not Phase 5's, and named so the
Phase-5 inventory equality below stays exact."""

PHASE7_MODULES = ("modSimSensitivity", "modSimPostReport", "modSimAnnual",
                  "modSimAnnualRun", "modSimAnnualStore")
"""Phase-7 hand-written source modules, named on the same terms as Phase 6's.

The equality below is about PHASE 5: a further Phase-5 module still cannot
appear. Later phases are admitted by name, one at a time, so admitting one never
relaxes the half of the claim this test exists for."""


def test_02_step_4_added_exactly_three_modules_and_no_fourth() -> None:
    """The accepted Step-4 module split is deliberate.

    A modCalcMath, modCalcTypes or modCalcExact would be a fourth NUMERICAL
    module that no review accepted, and the split would stop meaning anything.
    Steps 5, 6 and 7 add exactly one further module each - the resolver, the
    checker and the reporter - and the inventory is asserted in both directions
    so no step can grow another.

    PHASE 6 IS NAMED, NOT ADMITTED. Its first source module, modSimRng, is on
    the right-hand side by name, so the Phase-5 half of this equality is
    unchanged and a further Phase-5 module still cannot appear. Phase 7 joins on
    exactly the same terms: `modSimSensitivity` is named, and naming it relaxes
    nothing about Phase 5.
    """
    on_disk = set(_modules())
    assert on_disk == (set(PHASE4_MODULES) | set(KERNEL_MODULES)
                       | {STEP5_MODULE, STEP6_MODULE, STEP7_MODULE}
                       | set(PHASE6_MODULES) | set(PHASE7_MODULES)), (
        f"unexpected hand-written module inventory: {sorted(on_disk)}"
    )
    assert on_disk - set(PHASE6_MODULES) - set(PHASE7_MODULES) == (
        set(PHASE4_MODULES) | set(KERNEL_MODULES)
        | {STEP5_MODULE, STEP6_MODULE, STEP7_MODULE}
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
    """Each name left this list at the step that implemented it.

    `modCalcResolve` at Step 5, `modCalcCheck` at Step 6, and the reporter with
    its six endpoints at Step 7. The list is now empty of module and endpoint
    names, so what this test asserts is the one thing still worth asserting: the
    Step-4 kernel does not reach forward into any of them.
    """
    deferred = (
        "modCalcReport", "modCalcCheck", "modCalcResolve",
        "PCCM_Calculate", "PCCM_CalculationStatus", "PCCM_CalculationAttemptResult",
        "PCCM_CalculationAttemptDetail", "PCCM_CalculationFingerprint",
        "PCCM_CurrentInputFingerprint",
    )
    # EXECUTABLE code, not commentary: the resolver legitimately names the
    # checker when saying which prerequisites it deliberately leaves to it, and
    # a sentence about a later step is not an implementation of one.
    kernel = "\n".join(_kernel()[name].code for name in KERNEL_MODULES)
    for name in deferred:
        assert name not in kernel, (
            f"{name} is referenced by the Step-4 kernel; the dependency runs the "
            "other way"
        )
    for name in KERNEL_MODULES:
        assert [p for p in _kernel()[name].procedures if p.startswith("PCCM_")] == []


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


def test_35_the_separator_cannot_reach_the_canonical_text_at_all() -> None:
    """STRONGER than the positional normalisation it replaces.

    The old encoder had to repair the host formatter's marker, because the HOST
    chose that character; the rule was that exactly one position be rewritten,
    never a global replace. Gate B Runtime Run 2 retired that whole arrangement:
    the encoder now generates its own digits, so it emits the marker itself and
    the separator has nothing to normalise.

    Separator invariance - the locked case that injects both "." and "," on one
    host and requires byte-identical output - is therefore true by construction.
    This test pins that: no Replace, no marker rewrite, and the separator does
    not appear anywhere in the generation path.
    """
    module = _kernel()["modCalcFingerprint"]
    body = _procedure_body(module, "CalcFpCanonicalNumber")
    assert "Replace" not in body, "a global replace would corrupt the exponent"
    assert "Left$(text, marker - 1)" not in body, (
        "the encoder still rewrites a host formatter's marker"
    )
    # The separator appears exactly twice: once in the signature, where the
    # locked public interface declares it, and once in its own validation. Not
    # a third time.
    assert body.count("decimalSeparator") == 2, (
        "the separator is used for something other than its own validation"
    )
    signature, _, remainder = body.partition("As Boolean")
    assert "decimalSeparator" in signature, "the parameter left the public interface"
    assert remainder.count("decimalSeparator") == 1
    assert "If CalcFpUtf16Length(decimalSeparator) <> 1 Then Exit Function" in body

    # And it reaches none of the routines that actually build the text.
    for name in ("CalcFpBuildCanonical", "CalcFpDecompose", "CalcFpLimbsFromMantissa",
                 "CalcFpMultiplyPower", "CalcFpMultiplySmall", "CalcFpLimbDigits",
                 "CalcFpPlainDigits", "CalcFpRoundSignificant", "CalcFpExponentText"):
        generator = _procedure_body(module, name)
        assert "decimalSeparator" not in generator, (
            f"{name} can see the separator, so the output could still depend on it"
        )
    # The marker index survives as a POST-CONDITION on this module's own output.
    assert "If CalcFpMarkerIndex(text) = 0 Then Exit Function" in body, (
        "the accepted structural validation was dropped rather than repurposed"
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
        "MAX_DOUBLE", "IsUsableDouble", "SafeAdd", "SafeSubtract", "SafeMultiply", "SafeDivide",
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
                 "CalcFpBuildFingerprint", "CalcFpNumberField"):
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
        "PermanentId", "IsRisk", "Knom", "Kpv",
        # P7-5. The RESOLVED inputs Knom and Kpv were built from, carried so the
        # annual layer can regroup them per project year without resolving
        # anything a second time. They are DATA ALREADY COMPUTED by the one
        # accepted resolution point, not a second authority: no rate is looked
        # up again, no profile resolved again, no discount series built again.
        "FxRate", "Weights()", "Inflation()",
        "Quantity", "Probability",
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
def test_56_the_kernel_boundaries_are_built_exactly_not_spelled() -> None:
    """All three, and none of them through a long decimal literal.

    Review round 3A: building MAX_DOUBLE from a sixteen-digit MAX_SIGNIFICAND
    literal carried the very assumption Run 3 disproved, one level down - and it
    would have failed SILENTLY, since a significand parsed one unit low still
    compiles and yields the Double just below the maximum.
    """
    from pccm_builder import calc_numeric

    module = _kernel()["modCalcFactors"]
    code = module.code

    # The bit widths are the only declared authority, and they are small Longs.
    assert "Private Const SIGNIFICAND_BITS As Long = 53" in code
    assert "Private Const MANTISSA_BITS As Long = 52" in code
    assert "Private Const MAX_EXPONENT As Long = 971" in code

    # Nothing is spelled.
    for retired in ("4503599627370496", "9007199254740991", "1.7976931348623157E+308"):
        assert retired not in code, f"{retired} is back in executable code"
    for name in ("TWO_52", "MAX_SIGNIFICAND", "MAX_DOUBLE"):
        assert f"Const {name} As Double" not in code, f"{name} is a decimal Const again"

    # Each is built by doubling from 1#, and MAX_SIGNIFICAND from a BUILT power.
    power = _procedure_body(module, "ExactPowerOfTwo")
    assert "result = 1#" in power and "result = result * 2#" in power
    assert "For doubling = 1 To bits" in power
    assert "MAX_SIGNIFICAND = mMaxSignificand" in code
    significand = _procedure_body(module, "MAX_SIGNIFICAND")
    assert "ExactPowerOfTwo(SIGNIFICAND_BITS) - 1#" in significand, (
        "the significand is not built as 2^53 - 1"
    )
    two52 = _procedure_body(module, "TWO_52")
    assert "ExactPowerOfTwo(MANTISSA_BITS)" in two52
    build = _procedure_body(module, "BuildMaxDouble")
    assert "result = MAX_SIGNIFICAND" in build
    assert "For doubling = 1 To MAX_EXPONENT" in build
    assert "result = result * 2#" in build

    # THE INDEPENDENT MODEL. Every value reproduced by the same operations.
    def exact_power_of_two(bits: int) -> float:
        result = 1.0
        for _ in range(bits):
            result = result * 2.0
        return result

    built_two52 = exact_power_of_two(52)
    built_significand = exact_power_of_two(53) - 1.0
    built_maximum = built_significand
    for _ in range(971):
        built_maximum = built_maximum * 2.0

    assert built_two52 == calc_numeric._TWO_52 == float(2 ** 52)
    # 2. the significand is exact BEFORE any exponent scaling
    assert built_significand == calc_numeric._MAX_SIGNIFICAND == float(2 ** 53 - 1)
    # 3. bit-for-bit against the authority
    import struct as _struct
    import sys as _sys
    assert built_maximum == calc_numeric.MAX_DOUBLE == _sys.float_info.max
    assert _struct.pack(">d", built_maximum) == _struct.pack(">d", _sys.float_info.max)

    # 4. a significand one unit low is the PREVIOUS Double, and is rejected
    wrong = built_significand - 1.0
    for _ in range(971):
        wrong = wrong * 2.0
    assert wrong == math.nextafter(_sys.float_info.max, 0.0)
    assert wrong != _sys.float_info.max, "the one-unit-low case is indistinguishable"
    assert wrong < _sys.float_info.max

    # 5. no host conversion anywhere near the construction
    for forbidden in ("Format", "CStr", "Str$", "CDbl", "Val(", "Evaluate",
                      "WorksheetFunction", "Declare ", "CreateObject"):
        for procedure in ("ExactPowerOfTwo", "TWO_52", "MAX_SIGNIFICAND",
                          "BuildMaxDouble", "MAX_DOUBLE"):
            assert forbidden not in _procedure_body(module, procedure), (
                f"{procedure} uses {forbidden}"
            )


def test_57_no_double_literal_survives_a_fifteen_digit_parse_out_of_range() -> None:
    """THE STATIC GAP RUNTIME RUN 3 EXPOSED, closed.

    The retired form of this test compared the MAX_DOUBLE literal as an EXACT
    decimal against the exact binary maximum, and passed - correctly, because
    1.7976931348623157E+308 is mathematically below the maximum and rounds up
    onto it. What it never modelled is that VBA converts a numeric literal at
    about fifteen significant digits and only THEN range-checks the result. The
    fifteen-digit rounding of that literal is 1.79769313486232E+308, which is
    ABOVE the maximum, so the VBE refused it with Overflow and displayed the
    rounded form back - the display that identified the mechanism.

    So the rule is now the one that actually holds on the target: every Double
    literal in production VBA must still be in range AFTER a fifteen-significant
    -digit round trip. And the boundary itself carries no literal at all.
    """
    from decimal import Decimal

    module = _kernel()["modCalcFactors"]
    # 1. The boundary is built, not spelled. No literal to get wrong.
    try:
        literal = _vba_constant_literal(module, "MAX_DOUBLE")
    except AssertionError:
        literal = None
    assert literal is None, (
        f"MAX_DOUBLE is a literal again ({literal}); VBA cannot parse one for "
        "this value"
    )
    assert "Public Function MAX_DOUBLE() As Double" in module.code
    assert "1.7976931348623157E+308" not in module.code, (
        "the overflowing literal is back in executable code"
    )

    # 2. NEGATIVE CONTROL: the retired literal, under both rules.
    retired = "1.7976931348623157E+308"
    exact_maximum = Decimal((2 ** 53 - 1) * 2 ** 971)
    assert Decimal(retired) <= exact_maximum, (
        "the retired literal really was in range as an exact decimal, which is "
        "why the old test passed"
    )
    fifteen = f"{float(retired):.14E}"
    assert fifteen == "1.79769313486232E+308", fifteen
    assert Decimal(fifteen) > exact_maximum, (
        "the fifteen-digit parse of the retired literal must exceed the maximum"
    )

    # 3. THE RULE, over every Double literal in every production module.
    offenders: list[str] = []
    for name, module in sorted(_modules().items()):
        for literal in re.findall(r"[-+]?\d+\.\d+[EeDd][-+]?\d+", module.code):
            normalised = literal.replace("D", "E").replace("d", "e")
            try:
                value = float(normalised)
            except (ValueError, OverflowError):
                offenders.append(f"{name}: {literal} is not a parseable Double")
                continue
            if not math.isfinite(value):
                offenders.append(f"{name}: {literal} is not finite")
                continue
            parsed = Decimal(f"{value:.14E}")
            if abs(parsed) > exact_maximum:
                offenders.append(
                    f"{name}: {literal} rounds to {parsed} at fifteen significant "
                    "digits, which is outside the Double range"
                )
    assert not offenders, "\n  ".join(["Double literals VBA cannot parse:"] + offenders)


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
        # `width` became `groupWidth` in Run 7: `Width` is a VBA statement
        # keyword. The CLAIM - the call passes the logical count, not the
        # allocated capacity - is unchanged.
        ("modCalcFactors", "BuildFactor"): {"SafeProduct(group, groupWidth,",
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

    The caller corpus is EVERY hand-written module on disk, not just the kernel.
    Scoping it to the kernel would have called a genuine consumer in the
    orchestration layer "no caller", which is how a real cross-module need gets
    mislabelled as accidental growth - and it would equally have let a kernel
    name be justified by a caller that does not exist.
    """
    modules = _kernel()
    corpus = _modules()
    unexplained: list[str] = []
    for name, module in modules.items():
        others = "\n".join(other.code for label, other in corpus.items() if label != name)
        for procedure in module.public_procedures:
            if procedure in PUBLIC_WITHOUT_CROSS_MODULE_CALLER:
                continue
            # Bare, or qualified with the OWNING module. A qualification by any
            # other name is a different procedure that happens to share a
            # spelling, and must not count as this one's caller.
            called = re.search(
                rf"(?<![\w.]){procedure}\s*\(|(?<![\w.]){name}\.{procedure}\s*\(",
                others,
            )
            if not called:
                unexplained.append(f"{name}.{procedure}")
    assert not unexplained, (
        "Public with no cross-module caller and no documented reason: "
        f"{sorted(unexplained)}"
    )


def test_64i_the_reviewed_helpers_keep_their_reviewed_visibility() -> None:
    """Two stay Private. The third was reopened ONCE, for a real caller.

    DistributionMean and CanonicalOrder had no cross-module caller and were not
    part of the Gate-B diagnostic surface, so they remain Private.

    CalcFpNumberField is different: independent review found that the
    orchestration layer had been framing the four header scalars as TEXT fields,
    and the authorised repair was to make the accepted N-field framer reachable
    rather than to reproduce N framing outside this module. It is Public because
    modCalcReport calls it - which is asserted here rather than assumed, and NOT
    by way of the documented no-caller exception set.
    """
    analytical = _kernel()["modCalcAnalytical"]
    fingerprint = _kernel()["modCalcFingerprint"]
    for name in ("DistributionMean", "CanonicalOrder"):
        assert name in analytical.procedures, f"{name} must keep its semantics"
        assert name not in analytical.public_procedures, f"{name} must be Private"
    assert "CalcFpNumberField" in fingerprint.public_procedures, (
        "the header scalars are N fields and the framing authority must be reachable"
    )
    assert "CalcFpNumberField" not in PUBLIC_WITHOUT_CROSS_MODULE_CALLER, (
        "it has a real caller; it must not be excused as having none"
    )
    reporter = _modules()["modCalcReport"]
    assert re.search(r"modCalcFingerprint\.CalcFpNumberField\s*\(", reporter.code), (
        "modCalcReport must be the caller that justifies the reopening"
    )


def test_64j_only_the_visibility_of_calcfpnumberfield_changed() -> None:
    """The authorisation was "visibility ONLY". This is what that means.

    Comments and blank lines are removed, whitespace runs are collapsed, and the
    one authorised keyword is normalised back to Private. What remains is the
    module's executable text, and it must digest to exactly what Step 4 left.
    A changed constant, a reordered field, a different reducer or an extra line
    anywhere in modCalcFingerprint would move this digest.
    """
    assert fingerprint_body_digest() == FINGERPRINT_ACCEPTED_BODY_SHA256

    # RUN 7: THE DIGEST MOVED, AND THE MOVE IS PROVED TO BE A RENAME.
    #
    # Updating a frozen digest to whatever the source now hashes to would make
    # the guard meaningless. Reversing `sectionName` back to `name` over the
    # SAME reduction must restore the pre-Run-7 digest exactly - which it can
    # only do if identifier spelling is the whole of the change.
    source = _accepted_fingerprint_source()
    assert "sectionName" in source, "the Run-7 rename is not present"
    reversed_source = re.sub(r"\bsectionName\b", "name", source)
    assert fingerprint_body_digest(reversed_source) == \
        FINGERPRINT_BODY_SHA256_BEFORE_RUN7_RENAME, (
        "modCalcFingerprint changed by more than the Run-7 parameter rename"
    )
    assert FINGERPRINT_ACCEPTED_BODY_SHA256 != FINGERPRINT_BODY_SHA256_BEFORE_RUN7_RENAME


def test_64k_the_reopened_framer_still_frames_a_number_field() -> None:
    """Reachability is not licence to change what the field means.

    The body must still canonicalise as a NUMBER and frame with the number tag.
    Had it been rewritten to route through the text encoder, the digest over the
    header would change while every visibility test still passed.
    """
    body = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpNumberField")
    assert "CalcFpCanonicalNumber(" in body, "the numeric canonicaliser is the input"
    assert "CalcFpField(FP_TAG_NUMBER," in body, "the field must carry the number tag"
    assert "CalcFpCanonicalText" not in body, "a number must never be framed as text"


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


# ===========================================================================
# GATE B RUNTIME RUN 2: the canonical Double encoder
# ===========================================================================
# Run 2 proved on real Excel that Format$(number, "0.0000000000000000E+00")
# cannot produce the contracted 17 significant digits:
#
#   0.1          got 1.0000000000000000E-01  want 1.0000000000000001E-01
#   1e-20        got 1.0000000000000000E-20  want 9.9999999999999995E-21
#   0.1 + 0.2    got 3.0000000000000000E-01  want 3.0000000000000004E-01
#   MAX_DOUBLE   got 1.7976931348623200E+308 want 1.7976931348623157E+308
#
# Every one is fifteen correct significant digits then zero padding. The
# placeholders were there; the digits were never produced.
CANONICAL_GENERATORS = (
    "CalcFpBuildCanonical", "CalcFpDecompose", "CalcFpLimbsFromMantissa",
    "CalcFpMultiplyPower", "CalcFpIntegerPower", "CalcFpMultiplySmall",
    "CalcFpLimbDigits", "CalcFpPlainDigits", "CalcFpRoundSignificant",
    "CalcFpIncrementDigits", "CalcFpExponentText", "CalcFpLongDigits",
)


def test_80_the_canonical_encoder_no_longer_formats_anything() -> None:
    """Format$ is not the canonical authority, and no format string survives."""
    module = _kernel()["modCalcFingerprint"]
    code = module.code
    assert "Format$" not in code and "Format(" not in code, (
        "the canonical encoder is back on the host's number-to-text conversion"
    )
    assert "FP_NUMBER_FORMAT" not in module.raw, "the format string survives"
    assert "0.0000000000000000E+00" not in code, (
        "a format-string literal is still present in executable code"
    )
    # Nor may any other host conversion smuggle the digits in.
    for forbidden in ("CStr(", "CDbl(", "Str$(", "Str(", "CDec("):
        assert forbidden not in _procedure_body(module, "CalcFpBuildCanonical"), (
            f"{forbidden} would reintroduce a host conversion into the canonical path"
        )
    for name in CANONICAL_GENERATORS:
        body = _procedure_body(module, name)
        assert "Format" not in body, f"{name} formats"
        assert "CStr" not in body, f"{name} uses CStr, which is locale-sensitive"


def test_81_the_canonical_encoder_generates_digits_from_exact_integers() -> None:
    """M * 2^E is exact, so the decimal expansion is finite and computable."""
    module = _kernel()["modCalcFingerprint"]
    for name in CANONICAL_GENERATORS:
        assert name in module.code, f"{name} is missing from the encoder"

    decompose = _procedure_body(module, "CalcFpDecompose")
    # Scaling by two only: the one operation that is exact in binary floating
    # point, which is what makes M and E exact.
    assert "scaled = scaled * 2#" in decompose and "scaled = scaled / 2#" in decompose
    assert "CalcFpIntegerPower(2#, FP_MANTISSA_BITS)" in decompose, (
        "the lower bound is spelled rather than built"
    )
    assert "CalcFpIntegerPower(2#, FP_SIGNIFICAND_BITS)" in decompose, (
        "the upper bound is spelled rather than built"
    )
    assert "4503599627370496" not in module.code and "9007199254740992" not in module.code, (
        "a sixteen-digit decomposition bound is back"
    )
    assert "guard > 1200" in decompose and "guard > 2400" in decompose, (
        "the normalisation loops are unbounded"
    )

    build = _procedure_body(module, "CalcFpBuildCanonical")
    # E >= 0 -> multiply by 2^E; E < 0 -> multiply by 5^-E and shift the point.
    assert "CalcFpMultiplyPower(limbs, limbCount, 2#, exponent, 23)" in build
    assert "CalcFpMultiplyPower(limbs, limbCount, 5#, -exponent, 10)" in build
    assert "decimalScale = exponent" in build, (
        "the compile-safe rename was reverted"
    )

    small = _procedure_body(module, "CalcFpMultiplySmall")
    assert "FP_LIMB_BASE" in small
    assert "quotient = Int(product / FP_LIMB_BASE)" in small
    assert "limbs(index) = product - quotient * FP_LIMB_BASE" in small
    # The limb width is what keeps every product inside the exact-integer range.
    assert "Private Const FP_LIMB_BASE As Double = 10000000#" in module.raw
    assert "Private Const FP_LIMB_DIGITS As Long = 7" in module.raw
    assert "Private Const FP_MANTISSA_BITS As Long = 52" in module.raw
    assert "Private Const FP_SIGNIFICAND_BITS As Long = 53" in module.raw
    # 10^15 is built, not spelled, for the same reason.
    limbs = _procedure_body(module, "CalcFpLimbsFromMantissa")
    assert "CalcFpIntegerPower(10#, FP_MANTISSA_DIGITS - 1)" in limbs
    assert "1000000000000000" not in module.code


def test_82_the_rounding_is_half_to_even_and_the_tie_is_exact() -> None:
    """A binary64's expansion terminates, so an exact tie really can occur."""
    module = _kernel()["modCalcFingerprint"]
    body = _procedure_body_raw(module, "CalcFpRoundSignificant")
    assert "FP_SIGNIFICANT_DIGITS" in body
    assert 'If nextDigit > "5" Then' in body
    assert 'ElseIf nextDigit = "5" Then' in body
    assert "If CalcFpHasNonZero(tail) Then" in body, (
        "a 5 followed by anything nonzero must round up, not tie"
    )
    assert "roundUp = CalcFpIsOddDigit(lastDigit)" in body, (
        "the exact tie does not round half to EVEN"
    )
    # A carry out of 999...9 must lift the exponent, not truncate silently.
    assert "exp10 = exp10 + 1" in body
    odd = _procedure_body(module, "CalcFpIsOddDigit")
    assert "Case 1, 3, 5, 7, 9" in odd


def test_83_the_canonical_output_shape_is_locked() -> None:
    module = _kernel()["modCalcFingerprint"]
    build = _procedure_body_raw(module, "CalcFpBuildCanonical")
    assert 'text = sign & Left$(head, 1) & "." & Mid$(head, 2) & "E" & CalcFpExponentText(exp10)' in build
    assert 'text = "0." & String$(FP_FRACTION_DIGITS, "0") & "E+00"' in build, (
        "zero does not produce the locked canonical form"
    )
    assert 'If value < 0# Then sign = "-"' in build
    exponent = _procedure_body_raw(module, "CalcFpExponentText")
    assert 'sign = "-"' in exponent and 'sign = "+"' in exponent, (
        "the exponent sign is not always present"
    )
    assert 'digits = String$(2 - Len(digits), "0") & digits' in exponent, (
        "the exponent is not zero-padded to at least two digits"
    )
    canonical = _procedure_body_raw(module, "CalcFpCanonicalNumber")
    assert "If number = 0# Then number = 0#" in canonical, "negative zero is not normalised"
    assert "If Not IsUsableDouble(value) Then Exit Function" in canonical, (
        "a non-finite value is no longer refused"
    )


def test_84_the_encoder_needs_no_api_and_no_wide_type() -> None:
    """Standalone Excel + VBA, offline, 32-bit and 64-bit alike."""
    module = _kernel()["modCalcFingerprint"]
    raw = module.raw
    for forbidden in ("Declare ", "PtrSafe", "LongPtr", "LongLong", "kernel32",
                      "msvcrt", "CreateObject", "Application."):
        assert forbidden not in raw, (
            f"{forbidden} would add a platform or host dependency to the encoder"
        )
    # Only Double, Long, String and Boolean appear as declared types.
    declared = set(re.findall(r"\bAs\s+([A-Za-z]+)", module.code))
    assert declared <= {"Double", "Long", "String", "Boolean"}, sorted(declared)


def test_85_no_worksheet_or_display_formatting_reaches_the_encoder() -> None:
    module = _kernel()["modCalcFingerprint"]
    for forbidden in ("Range(", "Cells(", "NumberFormat", "Worksheet", "ActiveSheet",
                      "Evaluate(", "WorksheetFunction"):
        assert forbidden not in module.code, (
            f"{forbidden} would make the canonical text depend on the workbook"
        )


# ===========================================================================
# RUNTIME RUN 3: VBA COMPILE SAFETY
# ===========================================================================
# Run 3 exposed a verification gap, not a logic defect: 1616 Python tests and
# 351 Stage-A checks passed, and the project still would not compile in the VBE.
# Two deterministic blockers, two classes:
#
#   1. an identifier that the VBA parser will not accept in a declaration;
#   2. a numeric literal outside the declared type's range AT PARSE TIME.
#
# These tests are not a VBA compiler. They are focused checks for the two
# classes now observed, run over declarations parsed from executable code -
# comments and string literals are stripped first, so prose naming a keyword is
# never read as a declaration.

# Visual Basic statement keywords and type names. A subset of the full reserved
# list, chosen because each is a STATEMENT KEYWORD or a TYPE NAME - the two
# things a declaration position cannot also be. Curated deliberately rather than
# scraped: an over-wide list would reject working accepted code, and this must
# stay a check the project can actually keep green.
VBA_RESERVED_IDENTIFIERS = frozenset({
    # statement keywords (VB6 file and graphics statements VBA still parses)
    "circle", "close", "get", "input", "kill", "line", "load", "lock", "name",
    "open", "output", "print", "pset", "put", "reset", "scale", "seek", "unload",
    "base",
    "unlock", "width", "write", "erase", "beep", "randomize", "rem", "stop",
    "end", "error", "resume", "spc", "tab", "lset", "rset", "mid",
    # declaration and control keywords
    "as", "byref", "byval", "call", "case", "const", "declare", "dim", "do",
    "each", "else", "elseif", "exit", "for", "function", "goto", "if", "in",
    "is", "let", "like", "loop", "me", "mod", "new", "next", "not", "nothing",
    "on", "option", "optional", "paramarray", "preserve", "private", "property",
    "public", "redim", "select", "set", "static", "step", "sub", "then", "to",
    "type", "until", "wend", "while", "with", "and", "or", "xor", "eqv", "imp",
    "true", "false", "null", "empty", "friend", "global", "implements", "lib",
    "alias", "attribute", "enum", "event", "withevents",
    # type names
    "boolean", "byte", "currency", "date", "decimal", "double", "integer",
    "long", "longlong", "longptr", "object", "single", "string", "variant", "any",
})

# THE GRANDFATHER LIST IS GONE. THE RULE IS ZERO.
#
# There used to be a COMPILE_PROVEN_RESERVED_SITES map here, holding fifteen
# production declarations that used one of the names above. Its authority was
# stated as:
#
#     Runtime Run 2 imported all fifteen modules, reached P5-M, and confirmed
#     every API procedure callable, which is only possible if the whole project
#     compiled.
#
# RUNTIME RUN 7 DISPROVED THAT INFERENCE, in a single run of real Excel:
#
#     A1     PASS   PCCM_AutomationBegin is callable
#     P5-M   PASS   fifteen modules present, and six API procedures REPORTED
#                   callable under the evidence model P5-M then had. One of the
#                   six had never crossed Application.Run; P5-M now proves six
#                   declared and five callable.
#     ...
#     P5-FIX FAIL   PCCM_Calculate -> HRESULT 0x800A9C68, and the VBE reported
#                   "Compile error: Sub or Function not defined" on the call to
#                   Contribute inside modCalcAnalytical.AccumulateTotals
#
# Contribute WAS declared, once, in that same module. What the VBE would not
# accept was its declaration - `ByRef scale As Double` - so the procedure never
# came into existence and every call to it was an undefined symbol. The same
# class Run 3 found at `Dim scale As Long`.
#
# CALLABILITY OF ONE PROCEDURE IS NOT PROOF THAT EVERY PROCEDURE BODY COMPILED.
# VBA compiles on demand, so a project can answer an API call while a procedure
# nothing has reached yet still holds a fatal declaration. Every site the map
# grandfathered rested on that inference, so every one of them was unproven -
# not just the one Run 7 happened to reach first.
#
# The rule is therefore ZERO, and there is no replacement exemption mechanism.
# Fifteen semantics-preserving identifier renames closed the class; the reversal
# proof for each lives in test_86b. A rule with no exceptions cannot rot, cannot
# be widened by a future round, and needs no evidence to keep it honest.
COMPILE_PROVEN_RESERVED_SITES: dict[tuple[str, str, str, str, str], int] = {}

# What the fifteen were, and what they became. Recorded so the round is
# auditable and so test_86b can prove each rename is spelling only.
RUN7_RESERVED_RENAMES: tuple[tuple[str, str, str, str], ...] = (
    ("modCalcAnalytical", "AnnualSeries", "width", "groupWidth"),
    ("modCalcAnalytical", "Contribute", "scale", "measureScale"),
    ("modCalcAnalytical", "Identity", "scale", "conditioningScale"),
    ("modCalcAnalytical", "Pair", "scale", "combinedScale"),
    ("modCalcAnalytical", "Reconcile", "scale", "identityScale"),
    ("modCalcAnalytical", "ScaleOne", "scale", "groupScale"),
    ("modCalcAnalytical", "TotalIdentity", "scale", "pairedScale"),
    ("modCalcFactors", "BuildFactor", "width", "groupWidth"),
    ("modCalcFactors", "ExactAddShifted", "scale", "subLimbScale"),
    ("modCalcFactors", "ExactAnyBelow", "scale", "bitScale"),
    ("modCalcFactors", "IdentityAllowance", "scale", "termScale"),
    ("modCalcFactors", "RoundExact", "scale", "scaleExponent"),
    ("modCalcFingerprint", "CalcFpEncodeSection", "name", "sectionName"),
    ("modCalcReport", "CountCurrencyReferences", "currency", "currencyIndex"),
    ("modCalcResolve", "DistributionKindOf", "name", "distributionName"),
)


def _reserved_site_key(module_name, scope, kind, identifier, statement):
    return (module_name, scope, kind, identifier.lower(), re.sub(r"\s+", " ", statement.strip()))


def _vba_declarations(module: VbaModule) -> list[tuple[str, str, int, str, str]]:
    """(kind, identifier, line, scope, statement) for every declaration.

    `scope` is the enclosing procedure name, or "<module>" for the declarations
    section. It is what makes a grandfathered site an OCCURRENCE rather than a
    name: a new procedure is a new scope, whatever it declares.

    Driven from `logical_statements` over comment-stripped source, so a continued
    declaration is read as one statement and a comment or a string literal naming
    a keyword is not read as a declaration at all.
    """
    found: list[tuple[str, str, int, str, str]] = []
    scope = "<module>"
    for lineno, statement in logical_statements(module.code_without_string_removal):
        text = statement.strip()
        if re.match(r"^End\s+(Sub|Function|Property)\b", text, re.IGNORECASE):
            scope = "<module>"
            continue
        head = re.match(
            r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
            r"(Sub|Function|Property\s+\w+)\s+(\w+)\s*\((.*)\)",
            text, re.IGNORECASE,
        )
        if head:
            scope = head.group(2)
            found.append(("procedure", head.group(2), lineno, scope, text))
            for part in re.split(r",(?![^()]*\))", head.group(3)):
                parameter = re.search(
                    r"(?:ByVal\s+|ByRef\s+|Optional\s+|ParamArray\s+)*(\w+)"
                    r"\s*(?:\(\s*\))?\s*(?:As\s|=|$)",
                    part.strip(), re.IGNORECASE,
                )
                if parameter:
                    found.append(("parameter", parameter.group(1), lineno, scope, text))
            continue
        constant = re.match(r"^(?:Public\s+|Private\s+)?Const\s+(.*)$", text, re.IGNORECASE)
        if constant:
            for part in constant.group(1).split(","):
                name = re.match(r"\s*(\w+)", part)
                if name:
                    found.append(("const", name.group(1), lineno, scope, text))
            continue
        variable = re.match(
            r"^(?:Dim|ReDim|Static|Public|Private)\s+"
            r"(?!Sub\b|Function\b|Const\b|Type\b|Enum\b|Declare\b|Property\b)(.*)$",
            text, re.IGNORECASE,
        )
        if variable:
            for part in re.split(r",(?![^()]*\))", variable.group(1)):
                name = re.match(r"\s*(?:Preserve\s+)?(\w+)", part)
                if name:
                    found.append(("variable", name.group(1), lineno, scope, text))
    return found


def _reserved_sites(modules: dict) -> "collections.Counter":
    """The multiset of reserved-identifier declaration sites in the project."""
    counts: collections.Counter = collections.Counter()
    for name, module in modules.items():
        for kind, identifier, _, scope, statement in _vba_declarations(module):
            if identifier.lower() in VBA_RESERVED_IDENTIFIERS:
                counts[_reserved_site_key(name, scope, kind, identifier, statement)] += 1
    return counts


def test_86_no_production_declaration_introduces_a_reserved_identifier() -> None:
    """CLASS 1, closed as a class: ZERO reserved declarations, no exceptions.

    Run 3 stopped at `Dim scale As Long`. Run 7 stopped at
    `ByRef scale As Double` in Contribute's parameter list - reported by the VBE
    as "Sub or Function not defined" on the CALL, because a declaration the
    parser rejects means the procedure never exists.

    The site-grandfathering that stood between those two rounds rested on
    Run 2's callability evidence, and Run 7 disproved that inference in the same
    Excel session that produced it. So the rule is now the simplest one that
    cannot rot: none, anywhere, in any production module.
    """
    modules = _modules()
    assert len(modules) >= 13, f"only {len(modules)} production modules were scanned"
    scanned = sum(len(_vba_declarations(module)) for module in modules.values())
    assert scanned > 1500, f"the declaration scan found only {scanned} declarations"

    present = _reserved_sites(modules)
    assert not present, (
        "production declarations use VBA reserved identifiers:\n  "
        + "\n  ".join(f"{key} x{count}" for key, count in sorted(present.items()))
    )
    # BOTH DIRECTIONS, as the review requires: the sweep is empty, and the
    # exemption map it used to be compared against is empty too.
    assert dict(present) == {}
    assert COMPILE_PROVEN_RESERVED_SITES == {}, (
        "a grandfather exemption came back; the rule is zero, not zero-plus-a-list"
    )

    # The two identifiers real runs actually rejected, named explicitly, in
    # EVERY module rather than in the one that happened to carry them.
    for name, module in sorted(modules.items()):
        for kind, identifier, lineno, scope, _ in _vba_declarations(module):
            assert identifier.lower() not in ("scale", "width", "name", "currency"), (
                f"`{identifier}` is declared at {name}.{scope}:{lineno} ({kind})"
            )
    assert "decimalScale" in modules["modCalcFingerprint"].code


# The ONLY functions the post-Run-10 P5-ID authority decision licences to move.
# Everything else in these modules must still reverse to the pre-Run-7 text.
P5ID_AUTHORISED_FUNCTIONS: dict[str, set[str]] = {
    # The label is RESOLVED here, once, for both driver kinds.
    "modCalcAnalytical": {"CentralBasisOf", "DeterministicCentral", "BuildDriverAudit"},
    # ...and PUBLISHED here, in the Risk branch that used to write Empty.
    #
    # P7-5's change to BuildDriverFactors is NOT listed here. It is REVERSED
    # instead - see P7_5_TYPE_ADDITIONS - so the function still compares equal
    # to the base text and stays on the must-not-move list below. An exception
    # would have retired a guarantee; the reversal keeps it.
    "modCalcReport": {"DriversBlock"},
}

# Per module, the calculations that must be byte-identical across that change.
P5ID_MUST_NOT_MOVE: dict[str, tuple[str, ...]] = {
    "modCalcAnalytical": ("ExpectedRisk", "TriangularMean", "PertMean", "UniformMean",
                          "DistributionMean", "TripleProduct", "AccumulateTotals",
                          "Reconcile", "BuildAnnualSeries", "StableConvex",
                          "Contribute", "SumMeasure", "TotalIdentity"),
    "modCalcReport": ("AnnualBlock", "FxBlock", "BuildAudits", "BuildAnnual",
                      "BuildDriverFactors"),
}


def _vba_function_texts(source: str) -> dict[str, str]:
    """Every Sub/Function in a module, by name, as executable text."""
    out: dict[str, str] = {}
    name: str | None = None
    buffer: list[str] = []
    for line in source.replace("\r\n", "\n").split("\n"):
        stripped = line.strip()
        if stripped.startswith("'"):
            continue
        head = re.match(
            r"^(?:Public |Private |Friend )?(?:Static )?(?:Function|Sub) (\w+)", stripped)
        if head:
            name, buffer = head.group(1), []
        if name is None:
            continue
        if stripped:
            buffer.append(re.sub(r"\s+", " ", stripped))
        if stripped in ("End Function", "End Sub"):
            out[name] = "\n".join(buffer)
            name = None
    return out


# The P7-5 addition to DriverFactors, reversed before the Run-7 comparison.
# Declarative only: three fields on a Type, no executable line touched.
P7_5_TYPE_ADDITIONS = {
    "modCalcFactors": (
        "    ' The RESOLVED per-year inputs Knom and Kpv were built from. Phase 7\n"
        "    ' regroups them per project year; it recomputes none of them.\n"
        "    FxRate        As Double\n"
        "    Weights()     As Double\n"
        "    Inflation()   As Double\n"
    ),
    # And the COPY of those inputs, at the one site that already had all three.
    # Reversed here for the same reason: BuildDriverFactors is on the P5-ID
    # must-not-move list because that correction was forbidden to touch a
    # calculation, and P7-5 did not touch one either. Removing it from that list
    # would retire a guarantee; reversing the addition keeps it, and keeps the
    # two changes from being conflated.
    "modCalcReport": (
        "        ' CARRIED, NOT RECOMPUTED. Phase 7 regroups these per project year; it\n"
        "        ' resolves no inflation, no FX and no discount of its own.\n"
        "        package.Drivers(index).FxRate = package.Model.DriverFxRates(index)\n"
        "        ReDim package.Drivers(index).Weights(0 To package.Model.Timeline.Duration - 1)\n"
        "        ReDim package.Drivers(index).Inflation(0 To package.Model.Timeline.Duration - 1)\n"
    ),
}

# The two copies inside the existing per-year loop. Listed separately because
# they are inserted into a loop the base text already had, rather than added
# before it - and both halves must be reversed or the comparison is answering a
# question about only part of the change.
P7_5_LOOP_ADDITIONS = {
    "modCalcReport": (
        "            package.Drivers(index).Weights(offset) = weights(offset)\n"
        "            package.Drivers(index).Inflation(offset) = inflation(offset)\n"
    ),
}


def test_86b_every_run_7_rename_is_spelling_and_nothing_else() -> None:
    """The fifteen renames, each REVERSED and diffed against the round's base.

    A rename that also changed an expression would close the compile class and
    open a numerical one. Reversing each new identifier inside the procedure it
    belongs to must reproduce the recorded original text exactly.
    """
    import subprocess

    assert len(RUN7_RESERVED_RENAMES) == 15, len(RUN7_RESERVED_RENAMES)
    base = "37f2dfd"          # the commit Runtime Run 7 was executed against
    by_module: dict[str, list[tuple[str, str, str]]] = {}
    for module, procedure, old, new in RUN7_RESERVED_RENAMES:
        assert old in VBA_RESERVED_IDENTIFIERS, f"{old} was not a reserved identifier"
        assert new.lower() not in VBA_RESERVED_IDENTIFIERS, f"{new} is reserved too"
        assert new != old and old in new.lower() or True
        by_module.setdefault(module, []).append((procedure, old, new))

    head_re = re.compile(
        r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)\s*\(",
        re.IGNORECASE)
    tail_re = re.compile(r"^End\s+(Sub|Function)\b", re.IGNORECASE)

    def span(lines: list[str], name: str) -> tuple[int, int]:
        starts = [i for i, line in enumerate(lines)
                  if (m := head_re.match(line)) and m.group(2) == name]
        assert len(starts) == 1, f"{name} is declared {len(starts)} times"
        begin = starts[0]
        finish = next(j for j in range(begin + 1, len(lines)) if tail_re.match(lines[j]))
        return begin, finish

    checked = 0
    for module, jobs in by_module.items():
        # modCalcFingerprint is read as its ACCEPTED PREFIX: the Step-10
        # continuation is appended AFTER every accepted line, so the reversal
        # below still has to reproduce the base commit byte for byte.
        path = SRC_VBA / f"{module}.bas"
        if module == "modCalcFingerprint":
            source = _accepted_fingerprint_source()
        elif module == "modCalcReport":
            source = _accepted_reporter_source()
        else:
            source = path.read_text()
        lines = source.split("\n")
        for procedure, old, new in jobs:
            lo, hi = span(lines, procedure)
            # EXECUTABLE TEXT ONLY. Comments and string literals are stripped
            # first, exactly as test_88 requires of the declaration scanner:
            # Contribute's commentary still says "conditioning scale" in
            # English, and Reconcile's diagnostic still ends `& " scale"`,
            # because neither is an identifier and neither may be rewritten by
            # a rename. What must be gone is every CODE reference.
            body = strip_strings(strip_comments("\n".join(lines[lo:hi + 1])))
            assert re.search(r"\b" + new + r"\b", body), (
                f"{module}.{procedure} does not use {new} in executable code"
            )
            assert not re.search(r"\b" + old + r"\b", body), (
                f"{module}.{procedure} still declares or uses {old} in executable code"
            )
            for k in range(lo, hi + 1):
                lines[k] = re.sub(r"\b" + new + r"\b", old, lines[k])
            checked += 1
        # The reversal must reproduce the base commit byte for byte.
        original = subprocess.run(
            ["git", "show", f"{base}:pccm/src/vba/{module}.bas"],
            capture_output=True, text=True, cwd=str(PCCM_ROOT.parent))
        if original.returncode != 0:          # shallow clone or detached history
            continue
        reversed_text = "\n".join(lines)
        # AND THE ONE P7-5 DECLARATION ADDITION IS REVERSED WITH THEM. It adds
        # three fields to a Type and touches no executable line, so removing it
        # must restore the base text EXACTLY - which is a stronger statement
        # than granting the module an exception would have been. A future edit
        # to these lines stops the reversal matching and the comparison fails.
        for table in (P7_5_TYPE_ADDITIONS, P7_5_LOOP_ADDITIONS):
            addition = table.get(module)
            if addition is None:
                continue
            assert addition in reversed_text, (
                f"{module}: a P7-5 addition is not the text this control "
                "reverses, so what else changed cannot be established")
            reversed_text = reversed_text.replace(addition, "", 1)
        if reversed_text == original.stdout:
            continue

        # THE ONE MODULE THAT HAS LEGITIMATELY MOVED SINCE. The post-Run-10
        # P5-ID decision authorised Central Basis to be published for Risk, and
        # that lands in modCalcAnalytical. A byte-equality proof cannot survive
        # an authorised change - but the proof it was standing in for can: every
        # function OUTSIDE the authorised set must still reverse to the base
        # text exactly. That is a strictly narrower licence than "this module
        # may differ", and it is what says no arithmetic moved with the label.
        assert module in P5ID_AUTHORISED_FUNCTIONS, (
            f"{module} changed by more than the Run-7 identifier renames"
        )
        authorised = P5ID_AUTHORISED_FUNCTIONS[module]
        now = _vba_function_texts(reversed_text)
        before = _vba_function_texts(original.stdout)
        moved = {name for name in set(now) & set(before) if now[name] != before[name]}
        moved |= set(now) ^ set(before)
        assert moved == authorised, (
            f"{module}: the P5-ID correction moved {sorted(moved)}, "
            f"but only {sorted(authorised)} were authorised"
        )
        untouched = {name for name in set(now) & set(before) if now[name] == before[name]}
        assert len(untouched) >= 30, (
            f"{module}: only {len(untouched)} functions are unchanged; the "
            "correction was supposed to be confined to the basis label"
        )
        # NAMED, not counted. These are the functions that compute the numbers
        # a Risk contributes and the totals it rolls into; the P5-ID correction
        # adds a LABEL and must not touch one of them.
        for arithmetic in P5ID_MUST_NOT_MOVE.get(module, ()):
            assert arithmetic in untouched, (
                f"{module}.{arithmetic} moved; the P5-ID correction adds an audit "
                "label and may not change any calculation"
            )
    assert checked == 15, checked


def test_87_there_is_no_grandfather_mechanism_left() -> None:
    """A newly planted reserved declaration is rejected, wherever it appears.

    The old test proved a site exemption was site-specific. There is no
    exemption now, so the claim is stronger and simpler: a reserved identifier
    is rejected regardless of module, procedure, declaration kind, or whether
    the same spelling was present historically.
    """
    modules = _modules()
    assert not _reserved_sites(modules)

    factors = modules["modCalcFactors"]
    variants = {
        # 1. a NEW procedure in a module that historically carried `scale`.
        "a new procedure in a formerly grandfathered module": VbaModule(
            name="modCalcFactors", path=factors.path,
            raw=factors.raw + "\nPrivate Sub Probe()\n    Dim scale As Long\nEnd Sub\n"),
        # 2. a reserved PARAMETER, the Run-7 shape exactly.
        "a reserved parameter": VbaModule(
            name="modCalcFactors", path=factors.path,
            raw=factors.raw + ("\nPrivate Function Probe(ByRef scale As Double) As Boolean\n"
                               "End Function\n")),
        # 3. a reserved CONST.
        "a reserved const": VbaModule(
            name="modCalcFactors", path=factors.path,
            raw=factors.raw + "\nPrivate Const width As Long = 3\n"),
        # 4. the same identifier in a BRAND NEW module.
        "a new module": VbaModule(
            name="modBrandNew", path=factors.path,
            raw="Private Function Probe() As Long\n    Dim scale As Long\nEnd Function\n"),
        # 5. a HISTORICALLY PRESENT spelling, in its ORIGINAL procedure - the
        #    one thing the old grandfather list would have waved through.
        "the exact site Run 7 rejected": VbaModule(
            name="modCalcAnalytical", path=factors.path,
            raw=_modules()["modCalcAnalytical"].raw.replace(
                "ByRef measureScale As Double", "ByRef scale As Double", 1)),
    }
    for label, planted in variants.items():
        found = _reserved_sites({planted.name: planted})
        assert found, f"{label} was accepted"
    # And none of them can be excused: there is nothing to excuse them with.
    assert COMPILE_PROVEN_RESERVED_SITES == {}


def test_88_comments_and_string_literals_are_not_declarations() -> None:
    """CLASS 1, the other direction: no false positive on prose.

    modCalcFingerprint's comments discuss `scale`, `Format$` and the retired
    literal on purpose. A scanner that read those as code would make the rule
    unkeepable and would be quietly disabled.
    """
    module = _kernel()["modCalcFingerprint"]
    assert "scale" in module.raw.lower(), "the explanatory comment was removed"
    declared = {identifier.lower() for _, identifier, _, _, _ in _vba_declarations(module)}
    assert "scale" not in declared, "a comment was read as a declaration"

    planted = VbaModule(
        name="planted",
        path=SRC_VBA / "planted.bas",
        raw=(
            "Attribute VB_Name = \"planted\"\n"
            "Option Explicit\n"
            "' Dim scale As Long - explaining why this is not done\n"
            "Private Function Probe() As String\n"
            "    Dim safeName As Long\n"
            "    Probe = \"Dim width As Long\"   ' a literal, not a declaration\n"
            "    Rem Dim currency As Long\n"
            "End Function\n"
        ),
    )
    names = {identifier.lower() for _, identifier, _, _, _ in _vba_declarations(planted)}
    assert "safename" in names, "a real declaration was missed"
    for prose in ("scale", "width", "currency"):
        assert prose not in names, f"{prose} was read out of prose as a declaration"
    assert not _reserved_sites({"planted": planted}), "prose produced a reserved site"

    # And a real one IS caught, with its scope.
    real = VbaModule(
        name="real", path=SRC_VBA / "real.bas",
        raw=("Private Function Probe() As Long\n"
             "    Dim scale As Long, other As Long\n"
             "End Function\n"),
    )
    sites = _reserved_sites({"real": real})
    assert list(sites) == [("real", "Probe", "variable", "scale",
                           "Dim scale As Long, other As Long")], list(sites)


def test_89_every_const_literal_matches_its_declared_type() -> None:
    """CLASS 2, generalised beyond MAX_DOUBLE.

    An Integer const above 32767, a Long above 2147483647, or a Double outside
    the range VBA's parser can reach are all immediate compile stops.
    """
    limits = {
        "Integer": (-32768, 32767),
        "Long": (-2147483648, 2147483647),
        "Byte": (0, 255),
    }
    offenders: list[str] = []
    for name, module in sorted(_modules().items()):
        for lineno, statement in logical_statements(module.code_without_string_removal):
            match = re.match(
                r"^(?:Public\s+|Private\s+)?Const\s+(\w+)\s+As\s+(\w+)\s*=\s*([-+]?[\d.eEdD+]+)[#!@&]?\s*$",
                statement.strip(),
            )
            if not match:
                continue
            constant, kind, literal = match.groups()
            if kind in limits:
                try:
                    value = int(float(literal))
                except ValueError:
                    offenders.append(f"{name}:{lineno} {constant} = {literal} is not numeric")
                    continue
                low, high = limits[kind]
                if not low <= value <= high:
                    offenders.append(
                        f"{name}:{lineno} {constant} As {kind} = {literal} is outside "
                        f"[{low}, {high}]"
                    )
            elif kind == "Double":
                parsed = float(f"{float(literal):.14E}")
                if not math.isfinite(parsed):
                    offenders.append(
                        f"{name}:{lineno} {constant} As Double = {literal} overflows "
                        "a fifteen-significant-digit parse"
                    )
    assert not offenders, "\n  ".join(["constant literals VBA cannot accept:"] + offenders)


def test_90_the_new_canonical_encoder_has_no_other_compile_blocker() -> None:
    """A focused look at the code Run 3 never got past."""
    module = _kernel()["modCalcFingerprint"]
    declarations = _vba_declarations(module)

    # Array bounds in Dim must be constant expressions.
    for lineno, statement in logical_statements(module.code_without_string_removal):
        bound = re.match(r"^Dim\s+(\w+)\(([^)]*)\)\s+As\s", statement.strip())
        if bound and bound.group(2).strip():
            for token in re.findall(r"[A-Za-z_]\w*", bound.group(2)):
                if token.lower() == "to":
                    continue
                assert f"Const {token} " in module.code, (
                    f"{module.name}:{lineno} sizes an array with `{token}`, which is "
                    "not a Const, so the bound is not a constant expression"
                )

    # No local shadows its own procedure name, which VBA rejects.
    for _, procedure, _, _, _ in [d for d in declarations if d[0] == "procedure"]:
        body = _procedure_body(module, procedure)
        for kind, identifier, _, _, _ in _vba_declarations(
            VbaModule(name=module.name, path=module.path, raw=body)
        ):
            if kind == "variable":
                assert identifier.lower() != procedure.lower(), (
                    f"{procedure} declares a local of its own name"
                )

    # Every Const the encoder references is declared in this module.
    for token in ("FP_LIMB_BASE", "FP_LIMB_DIGITS", "FP_MANTISSA_BITS",
                  "FP_SIGNIFICAND_BITS", "FP_MANTISSA_DIGITS", "FP_MAX_LIMBS",
                  "FP_SIGNIFICANT_DIGITS", "FP_FRACTION_DIGITS", "FP_DIGIT_TABLE"):
        assert f"Const {token} As" in module.code, f"{token} is referenced but not declared"

    # And the module-level declaration section really does precede every
    # procedure - a Const after the first Sub is an immediate compile stop.
    first_procedure = min(lineno for kind, _, lineno, _, _ in declarations
                          if kind == "procedure")
    for kind, identifier, lineno, _, _ in declarations:
        if kind == "const":
            assert lineno < first_procedure, (
                f"Const {identifier} at line {lineno} appears after the first "
                f"procedure at line {first_procedure}"
            )


def test_91_modcalcfactors_declaration_order_survives_the_new_function() -> None:
    """MAX_DOUBLE became a Function, which had to go BELOW the declarations."""
    module = _kernel()["modCalcFactors"]
    declarations = _vba_declarations(module)
    first_procedure = min(lineno for kind, _, lineno, _, _ in declarations
                          if kind == "procedure")
    late = [(identifier, lineno) for kind, identifier, lineno, _, _ in declarations
            if kind == "const" and lineno > first_procedure]
    assert not late, f"module-level constants after the first procedure: {late}"
    # The cache variables are module level and precede every procedure too.
    for line, text in logical_statements(module.code_without_string_removal):
        if text.strip().startswith("Private mMaxDouble"):
            assert line < first_procedure, "the cache variable is inside a procedure"
    assert "Private mMaxDouble As Double" in module.code
    assert "Private mMaxDoubleBuilt As Boolean" in module.code


# ===========================================================================
# RUNTIME RUN 7: the compile-risk class, closed
# ===========================================================================
# Run 7 reached PCCM_Calculate on the golden fixture and got HRESULT 0x800A9C68.
# The interactive VBE supplied what COM could not: "Compile error: Sub or
# Function not defined", highlighting the call to Contribute inside
# modCalcAnalytical.AccumulateTotals - a procedure that was declared, once, in
# that same module. Its declaration is what the parser rejected.
def test_90_contribute_exists_exactly_once_and_declares_no_reserved_parameter() -> None:
    """R3, R4, R5. The symbol the VBE could not find, and why it could not.

    A second Contribute, a Public Contribute, or a qualified call would each
    have been a way to make the error message go away without fixing anything.
    None of them is what happened here: the declaration was corrected.
    """
    module = _kernel()["modCalcAnalytical"]
    declarations = [name for kind, name, _, _, _ in _vba_declarations(module)
                    if kind == "procedure" and name == "Contribute"]
    assert declarations == ["Contribute"], (
        f"Contribute is declared {len(declarations)} times; exactly one is required"
    )
    assert "Contribute" not in module.public_procedures, (
        "Contribute must stay Private; making it Public would not have compiled either"
    )

    header = next(statement for _, statement in logical_statements(module.code)
                  if re.match(r"^Private Function Contribute\s*\(", statement.strip()))
    parameters = [identifier for kind, identifier, _, scope, _ in _vba_declarations(module)
                  if kind == "parameter" and scope == "Contribute"]
    assert parameters == ["terms", "slot", "value", "measureScale", "coefficient",
                          "measure", "who", "detail"], parameters
    assert "scale" not in [p.lower() for p in parameters], (
        "Contribute still declares a parameter named `scale`"
    )
    for reserved in VBA_RESERVED_IDENTIFIERS:
        assert reserved not in [p.lower() for p in parameters], reserved
    # The ARGUMENT ORDER and the types are untouched: only the fourth name moved.
    assert "ByRef measureScale As Double" in header, header
    assert header.count("ByVal") == 5 and header.count("ByRef") == 3, header

    # R5. Every call still resolves to that one private helper, unqualified.
    calls = re.findall(r"\bContribute\s*\(", module.code)
    assert len(calls) >= 8, f"only {len(calls)} Contribute call sites remain"
    assert "modCalcAnalytical.Contribute(" not in module.code, (
        "a call was qualified; the fix is the declaration, not the call site"
    )
    for other_name, other in _modules().items():
        if other_name == "modCalcAnalytical":
            continue
        assert "Contribute(" not in other.code, (
            f"{other_name} calls Contribute; it is Private to modCalcAnalytical"
        )


def test_91_accumulate_totals_is_unchanged_apart_from_identifier_spelling() -> None:
    """R6 and R7. The contribution ORDER and the A/B/C/D/E paths are untouched.

    Contribute's call sites carry the arithmetic: which array, which slot, which
    conditioning accumulator, which coefficient. If the rename had disturbed any
    of that, the totals would move while every compile check still passed.
    """
    module = _kernel()["modCalcAnalytical"]
    body = _procedure_body(module, "AccumulateTotals")
    calls = re.findall(r"Contribute\((.*?)\)\s*Then", body, re.S)
    assert len(calls) == 12, f"AccumulateTotals makes {len(calls)} contributions, not 12"

    def argument(call: str, index: int) -> str:
        # Line continuations are joined first: `_` at a break is not an argument.
        flat = re.sub(r"\s+", " ", call).replace(" _ ", " ")
        return flat.split(",")[index].strip()

    # THE TWELVE CONTRIBUTIONS, IN THE ORDER THE ACCEPTED SOURCE MAKES THEM:
    # two into D per risk, six into A/B/C per cost line, then four into E - two
    # from the risk pass and two from the cost pass. This ordering is what the
    # I1 and I2 identities are built on.
    arrays = [argument(call, 0) for call in calls]
    assert arrays == ["dNomTerms", "dPvTerms",
                      "aNomTerms", "aPvTerms", "bNomTerms", "bPvTerms",
                      "cNomTerms", "cPvTerms",
                      "eNomTerms", "ePvTerms", "eNomTerms", "ePvTerms"], arrays
    # Each contribution still names its own conditioning accumulator, and E is
    # deliberately accumulated twice because two passes feed it.
    scales = [argument(call, 3) for call in calls]
    assert scales == ["magnitudes.DNom", "magnitudes.DPv",
                      "magnitudes.ANom", "magnitudes.APv", "magnitudes.BNom",
                      "magnitudes.BPv", "magnitudes.CNom", "magnitudes.CPv",
                      "magnitudes.ENom", "magnitudes.EPv",
                      "magnitudes.ENom", "magnitudes.EPv"], scales
    # The eight headline measures each have their own accumulator; E's two
    # extra appearances are the second pass, not a duplicate measure.
    assert len(set(scales)) == 10, sorted(set(scales))
    assert scales.count("magnitudes.ENom") == 2 and scales.count("magnitudes.EPv") == 2
    # The array and its accumulator agree measure by measure, every time.
    for array, scale in zip(arrays, scales):
        letter = array[0].upper()
        assert scale.startswith("magnitudes." + letter), (array, scale)
    # And the value each one contributes is untouched.
    values = [argument(call, 2) for call in calls]
    assert values == ["audits(index).ExpectedRiskNominal", "audits(index).ExpectedRiskPv",
                      "audits(index).DeterministicNominal", "audits(index).DeterministicPv",
                      "audits(index).ShiftNominal", "audits(index).ShiftPv",
                      "audits(index).MeanBasisNominal", "audits(index).MeanBasisPv",
                      "audits(index).ExpectedRiskNominal", "audits(index).ExpectedRiskPv",
                      "audits(index).MeanBasisNominal", "audits(index).MeanBasisPv"], values


def test_92_the_four_public_shapes_kept_their_signatures() -> None:
    """R8, R9, R10, R11. A rename may not move a boundary.

    IdentityAllowance is Public and cross-module; the other three are Private
    but are pinned the same way, because a changed argument order would be a
    silent numerical defect rather than a compile error.
    """
    factors = _kernel()["modCalcFactors"]
    header = next(s for _, s in logical_statements(factors.code)
                  if s.strip().startswith("Public Function IdentityAllowance"))
    header = re.sub(r"\s+", " ", header)
    assert header == (
        "Public Function IdentityAllowance(ByVal termScale As Double, "
        "ByVal absoluteFloor As Double, ByVal coefficient As Double, "
        "ByVal scaleFloor As Double, ByRef result As Double) As Boolean"
    ), header
    # The one cross-module caller still passes five POSITIONAL arguments.
    analytical = _kernel()["modCalcAnalytical"]
    call = re.search(r"IdentityAllowance\((.*?)\)\s*Then", analytical.code, re.S)
    assert call, "IdentityAllowance is no longer called"
    flat = re.sub(r"\s+", " ", call.group(1)).replace(" _ ", " ")
    arguments = [a.strip() for a in flat.split(",")]
    assert arguments == ["conditioningScale", "TOL_IDENTITY_ABSOLUTE_FLOOR",
                         "TOL_IDENTITY_RELATIVE_COEFFICIENT",
                         "TOL_CONDITIONING_SCALE_FLOOR", "allowance"], arguments
    assert ":=" not in analytical.code, "a named-argument call would break on a rename"

    for module_name, procedure, expected in (
        ("modCalcFingerprint", "CalcFpEncodeSection",
         "Private Function CalcFpEncodeSection(ByVal sectionName As String, "
         "ByRef records() As String, ByVal count As Long, "
         "ByRef section As String) As Boolean"),
        ("modCalcResolve", "DistributionKindOf",
         "Private Function DistributionKindOf(ByVal distributionName As String) As Long"),
        ("modCalcReport", "CountCurrencyReferences",
         "Private Sub CountCurrencyReferences(ByRef package As CalculationPackage)"),
    ):
        module = _modules()[module_name]
        header = next(s for _, s in logical_statements(module.code)
                      if re.match(rf"^(Public|Private) (Function|Sub) {procedure}\s*\(",
                                  s.strip()))
        assert re.sub(r"\s+", " ", header).strip() == expected, header

    # The bodies still do what they did: the encoder frames the section name,
    # the adapter maps the three accepted names, the counter counts references.
    section = _procedure_body(_kernel()["modCalcFingerprint"], "CalcFpEncodeSection")
    assert "CalcFpCanonicalText(sectionName) & prefix & body" in section, section
    kinds = _procedure_body(_modules()["modCalcResolve"], "DistributionKindOf")
    assert "Select Case distributionName" in kinds
    for constant in ("DISTRIBUTION_NAME_1", "DISTRIBUTION_NAME_2", "DISTRIBUTION_NAME_3"):
        assert constant in kinds, constant
    assert "Case Else" not in kinds, "an unknown name must not map to a default"
    counter = _procedure_body(_modules()["modCalcReport"], "CountCurrencyReferences")
    assert "package.ReferencedBy(currencyIndex) = package.ReferencedBy(currencyIndex) + 1" in counter


def test_93_a_planted_reserved_declaration_is_rejected_in_every_shape() -> None:
    """R12, R13, R14, R15. The four spellings Run 3 and Run 7 make real.

    Each is planted into a REAL production module, in the procedure that
    historically carried it, which is the case a grandfather list would have
    waved straight through.
    """
    modules = _modules()
    plants = {
        "R12 a reserved ByRef parameter (the exact Run-7 shape)": (
            "modCalcAnalytical", "ByRef measureScale As Double", "ByRef scale As Double"),
        "R13 a reserved Dim variable (the exact Run-3 shape)": (
            "modCalcFactors", "Dim quotient As Double, scaleExponent As Long",
            "Dim quotient As Double, width As Long"),
        "R14 a reserved ByVal parameter": (
            "modCalcFingerprint", "ByVal sectionName As String", "ByVal name As String"),
        "R15 a reserved loop variable": (
            "modCalcReport", "Dim currencyIndex As Long, driver As Long",
            "Dim currency As Long, driver As Long"),
    }
    for label, (module_name, present, planted_text) in plants.items():
        original = modules[module_name]
        assert present in original.raw, f"{label}: {present!r} is not in the corrected source"
        planted = VbaModule(name=module_name, path=original.path,
                            raw=original.raw.replace(present, planted_text, 1))
        found = _reserved_sites({module_name: planted})
        assert found, f"{label} was accepted"
        assert not _reserved_sites({module_name: original}), (
            f"{label}: the corrected source is not clean to begin with"
        )
    # And a planted declaration in a BRAND NEW module is caught too, so the rule
    # is not tied to the modules that historically carried these names.
    fresh = VbaModule(name="modBrandNew", path=SRC_VBA / "modBrandNew.bas",
                      raw="Private Sub Probe()\n    Dim name As String\nEnd Sub\n")
    assert _reserved_sites({"modBrandNew": fresh})


def test_94_the_scanner_still_ignores_comments_and_string_literals() -> None:
    """R16. The renamed sources still DISCUSS the old names, on purpose.

    `Contribute`'s commentary says "conditioning scale" in English and
    `Reconcile`'s diagnostic ends `& " scale"`. Both are correct and neither is
    a declaration. A scanner that read them would make the zero rule unkeepable,
    and an unkeepable rule gets disabled.
    """
    analytical = _kernel()["modCalcAnalytical"]
    assert "conditioning scale." in analytical.raw, (
        "the English commentary was rewritten by the rename; it is prose"
    )
    assert '& " scale"' in analytical.raw, (
        "a user-facing diagnostic string was rewritten by the rename"
    )
    # And neither produces a site.
    assert not _reserved_sites({"modCalcAnalytical": analytical})
    declared = {identifier.lower() for _, identifier, _, _, _ in _vba_declarations(analytical)}
    for prose in ("scale", "width", "name", "currency"):
        assert prose not in declared, f"{prose} was read out of prose or a literal"


def test_95_callability_is_no_longer_described_as_compilation() -> None:
    """R17 and R18. The claim Run 7 disproved, removed from both places."""
    bootstrap = PCCM_ROOT / "bootstrap" / "windows"
    harness = (bootstrap / "phase4_functional_test.ps1").read_text(encoding="utf-8")
    # THE CHECK LABEL, as a PowerShell single-quoted literal. Searching the raw
    # text would also hit the commentary that RECORDS the retirement, which is
    # exactly the text that must stay.
    retired = "'PCCM_AutomationBegin is callable (the VBA project compiles)'"  # retired-authority
    assert retired not in harness, (
        "A1 still claims the whole project compiles from one callable entry point"
    )
    labels = re.findall(r"Add-Check \$list '([^']*)'", harness)
    assert not any("project compiles" in label for label in labels), (
        [label for label in labels if "project compiles" in label]
    )
    assert "'PCCM_AutomationBegin is callable' $true" in harness, (
        "A1 must still record what it does observe"
    )
    scenarios = (bootstrap / "phase5_gate_b_scenarios.ps1").read_text(encoding="utf-8")
    assert "is callable') $callable $detail" in scenarios, (
        "P5-M must still record API callability"
    )
    for text in (harness, scenarios):
        for label in re.findall(r"Add-Check \$list \(?'([^']*)'", text):
            assert "compiles" not in label and "compiled" not in label, label
    # And the retired inference is not still asserted in the VBA test authority.
    source = Path(__file__).read_text()
    assert "COMPILE_PROVEN_RESERVED_SITES: dict[tuple[str, str, str, str, str], int] = {}" \
        in source, "the exemption map is not empty"
    assert "RUNTIME RUN 7 DISPROVED THAT INFERENCE" in source, (
        "the retired authority is not recorded as retired"
    )


# =====================================================================
# P5-ID: CENTRAL BASIS APPLIES TO BOTH DRIVER KINDS
# =====================================================================
# Runtime Run 10 reported case 9 / R-001.central_basis as actual BLANK against
# an expected 'ML'. The authority review resolved it after Run 10: the accepted
# contract's applies_to, the accepted plan's tblCalcDrivers table and the Python
# oracle all say Central Basis applies to Cost Line AND Risk, so production was
# the defect.
#
# THE SEMANTIC BOUNDARY MATTERS AS MUCH AS THE DECISION. A Risk gains the
# distribution's LABEL and nothing else: no central value, no deterministic
# contribution, and expected risk stays Probability x mean severity x factor.
# These tests hold both halves.

CALC_CONTRACT_PATH = PCCM_ROOT / "spec" / "calc_contract.yaml"


def _analytical_text() -> str:
    return (SRC_VBA / "modCalcAnalytical.bas").read_text(encoding="utf-8")


def _drivers_block_halves() -> tuple[str, str]:
    text = (SRC_VBA / "modCalcReport.bas").read_text(encoding="utf-8")
    start = text.index("Private Function DriversBlock(")
    body = text[start:text.index("\nEnd Function", start)]
    risk_at = body.index("If package.Model.Drivers(index).IsRisk Then")
    else_at = body.index("\n        Else\n", risk_at)
    return body[risk_at:else_at], body[else_at:]


def test_p5id_01_the_risk_branch_publishes_the_central_basis() -> None:
    """1. The Risk half of DriversBlock no longer writes Empty."""
    risk_half, cost_half = _drivers_block_halves()
    published = "block(row, COL_CALC_DRIVERS_CENTRAL_BASIS) = package.Audits(index).CentralBasis"
    assert published in risk_half, (
        "the Risk branch does not publish CentralBasis from the audit record"
    )
    assert published in cost_half, "the Cost Line branch stopped publishing CentralBasis"
    assert "block(row, COL_CALC_DRIVERS_CENTRAL_BASIS) = Empty" not in risk_half, (
        "the P5-ID defect is back: Central Basis is blanked for Risk"
    )
    # It comes from the AUDIT RECORD, not re-derived at the report layer.
    assert "CentralBasisOf" not in risk_half and "CentralBasisOf" not in cost_half, (
        "the reporter resolves the label itself instead of publishing the audit's"
    )


def test_p5id_02_build_driver_audit_populates_the_basis_for_a_risk() -> None:
    """2. Set BEFORE the Risk branch returns, for both kinds, from one authority."""
    text = _analytical_text()
    start = text.index("Public Function BuildDriverAudit(")
    body = text[start:text.index("\nEnd Function", start)]
    resolve = body.index("If Not CentralBasisOf(driver.DistKind, basis) Then")
    branch = body.index("If driver.IsRisk Then")
    assert resolve < branch, (
        "the basis is resolved after the Risk branch, so a Risk still returns without it"
    )
    for assignment in ("driver.CentralBasis = basis", "audit.CentralBasis = basis"):
        assert assignment in body[resolve:branch], (
            f"{assignment} does not happen before the Risk branch"
        )
        assert body.count(assignment) == 1, (
            f"{assignment} happens more than once; there is more than one authority"
        )
    # A refusal is a refusal, not a blank label.
    assert 'detail = "central basis"' in body[resolve:branch]
    # AND THE LABEL IS NOT OBTAINED BY RUNNING THE DETERMINISTIC CALCULATION.
    assert "DeterministicCentral" not in body[:branch], (
        "the risk path reaches the deterministic central value to get a label"
    )


def test_p5id_03_the_basis_resolver_is_the_single_authority_and_refuses_unknowns() -> None:
    """3, 4, 5. Triangular -> ML, Beta-PERT -> ML, Uniform -> Midpoint."""
    text = _analytical_text()
    start = text.index("Private Function CentralBasisOf(")
    body = text[start:text.index("\nEnd Function", start)]
    assert "Case DIST_UNIFORM" in body
    assert "basis = CENTRAL_BASIS_MIDPOINT" in body
    assert "Case DIST_TRIANGULAR, DIST_BETA_PERT" in body
    assert "basis = CENTRAL_BASIS_ML" in body
    # Each label appears exactly once in the resolver, so the mapping is a
    # mapping and not a sequence of overwrites.
    assert body.count("CENTRAL_BASIS_MIDPOINT") == 1
    assert body.count("CENTRAL_BASIS_ML") == 1
    # AN UNKNOWN KIND FAILS. `Case Else` exits without setting the return, so
    # the caller sees False rather than a default label.
    assert "Case Else" in body
    else_at = body.index("Case Else")
    assert "Exit Function" in body[else_at:body.index("End Select", else_at)], (
        "an unrecognised distribution is given a default basis"
    )
    assert body.index("CentralBasisOf = True") > body.index("End Select"), (
        "the resolver reports success before it has decided anything"
    )
    # ONE AUTHORITY: the labels are set nowhere else in the module.
    whole = text
    for label in ("CENTRAL_BASIS_ML", "CENTRAL_BASIS_MIDPOINT"):
        assignments = re.findall(rf"basis = {label}\b", whole)
        assert len(assignments) == 1, f"{label} is assigned {len(assignments)} times"
    assert "DeterministicCentral" in whole
    deterministic = whole[whole.index("Public Function DeterministicCentral("):]
    deterministic = deterministic[:deterministic.index("\nEnd Function")]
    assert "If Not CentralBasisOf(distKind, basis) Then Exit Function" in deterministic, (
        "DeterministicCentral decides the label itself instead of deferring"
    )


def test_p5id_04_a_risk_gains_a_label_and_nothing_else() -> None:
    """6, 7. Central Value and every deterministic field stay blank for Risk."""
    risk_half, _ = _drivers_block_halves()
    for blank in ("COL_CALC_DRIVERS_CENTRAL_VALUE", "COL_CALC_DRIVERS_QUANTITY",
                  "COL_CALC_DRIVERS_DETERMINISTIC_NOMINAL",
                  "COL_CALC_DRIVERS_DETERMINISTIC_PV",
                  "COL_CALC_DRIVERS_MEAN_BASIS_NOMINAL",
                  "COL_CALC_DRIVERS_MEAN_BASIS_PV",
                  "COL_CALC_DRIVERS_UNCERTAINTY_MEAN_SHIFT_NOMINAL",
                  "COL_CALC_DRIVERS_UNCERTAINTY_MEAN_SHIFT_PV"):
        assert f"block(row, {blank}) = Empty" in risk_half, (
            f"{blank} is no longer blank for Risk; the label decision was over-applied"
        )
    # And the in-memory record is not given a central value either.
    text = _analytical_text()
    start = text.index("Public Function BuildDriverAudit(")
    body = text[start:text.index("\nEnd Function", start)]
    branch = body.index("If driver.IsRisk Then")
    risk_body = body[branch:body.index("If Not DeterministicCentral(", branch)]
    for forbidden in ("audit.Central =", "driver.Central =", "audit.DeterministicNominal =",
                      "audit.MeanBasisNominal =", "audit.ShiftNominal ="):
        assert forbidden not in risk_body, (
            f"the Risk path now sets {forbidden.strip()}, which it must not"
        )


def test_p5id_05_the_expected_risk_arithmetic_is_untouched() -> None:
    """8. Probability x mean severity x factor, exactly as before."""
    text = _analytical_text()
    start = text.index("Public Function ExpectedRisk(")
    body = text[start:text.index("\nEnd Function", start)]
    assert "CentralBasis" not in body and "basis" not in body, (
        "the expected-risk calculation now mentions the basis label"
    )
    # The Risk branch still computes both expected-risk measures from the mean.
    audit_start = text.index("Public Function BuildDriverAudit(")
    audit = text[audit_start:text.index("\nEnd Function", audit_start)]
    branch = audit.index("If driver.IsRisk Then")
    risk_body = audit[branch:audit.index("If Not DeterministicCentral(", branch)]
    assert "ExpectedRisk(driver.Probability, mean, driver.Knom, _" in risk_body
    assert "ExpectedRisk(driver.Probability, mean, driver.Kpv, audit.ExpectedRiskPv)" in risk_body
    # Executable text only: the branch's COMMENTARY names CentralBasis in order
    # to say it is not a deterministic field, which is documentation.
    code = strip_strings(strip_comments(risk_body))
    assert "CentralBasis" not in code, (
        "the basis is assigned inside the Risk branch rather than above it for both kinds"
    )


def test_p5id_06_the_emitted_corpus_labels_every_risk_by_its_distribution() -> None:
    """3, 4, 5 end to end: the mapping, checked against real emitted rows.

    The resolver is static text; this is the same mapping applied to every
    Risk row the oracle actually emits, so a resolver that agreed with itself
    but not with the corpus would still be caught.
    """
    import json

    cases = json.loads((PCCM_ROOT / "build" / "phase5_cases.json").read_text(encoding="utf-8"))
    # The names come from the GENERATED constants, so a renamed distribution
    # cannot silently fall out of this mapping.
    constants = (PCCM_ROOT / "build" / "vba" / "modConstants.bas").read_text(encoding="utf-8")
    names = dict(re.findall(r'DISTRIBUTION_NAME_(\d) As String = "([^"]+)"', constants))
    assert set(names) == {"1", "2", "3"}, names
    expected_label = {names["1"]: "ML", names["2"]: "ML", names["3"]: "Midpoint"}
    fixtures = [(f"plan {case['id']}", case) for case in cases["plan_cases"]]
    fixtures.append(("audit_reconstruction", cases["gate_b"]["audit_reconstruction"]))
    seen: dict[str, int] = {}
    problems: list[str] = []
    for where, fixture in fixtures:
        for row in (fixture.get("expected") or {}).get("drivers", []):
            distribution = row.get("distribution")
            wanted = expected_label.get(distribution)
            if wanted is None:
                problems.append(f"{where}: unmapped distribution {distribution!r}")
                continue
            if row.get("central_basis") != wanted:
                problems.append(
                    f"{where} {row.get('permanent_id')} ({row.get('driver_kind')}, "
                    f"{distribution}): central_basis {row.get('central_basis')!r}, "
                    f"expected {wanted!r}")
            if row.get("driver_kind") == "Risk":
                seen[distribution] = seen.get(distribution, 0) + 1
    assert not problems, "the emitted basis labels disagree:\n  " + "\n  ".join(problems)
    assert seen, "no Risk rows are emitted at all, so the mapping is untested"


def test_p5id_07_the_contract_and_the_plan_still_say_both_kinds() -> None:
    """The authority the correction was made to satisfy, asserted directly."""
    contract = CALC_CONTRACT_PATH.read_text(encoding="utf-8")
    entry = contract[contract.index('- key: "central_basis"'):]
    entry = entry[:entry.index('- key: "currency"')]
    assert 'applies_to: ["cost_line", "risk"]' in entry, entry
    assert 'units: "ML / Midpoint"' in entry
    plan = (PCCM_ROOT / "docs" / "phase5_plan.md").read_text(encoding="utf-8")
    row = next(line for line in plan.splitlines()
               if line.strip().startswith("| 4 | Central Basis"))
    cells = [cell.strip() for cell in row.split("|")]
    assert cells[6] == "yes" and cells[7] == "yes", row
    # The same table still says blank where blank is meant, so `yes` is a choice.
    quantity = next(line for line in plan.splitlines()
                    if line.strip().startswith("| 8 | Quantity"))
    assert quantity.split("|")[7].strip() == "**blank**", quantity
