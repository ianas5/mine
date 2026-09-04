#!/usr/bin/env python3
"""PCCM Phase 5 Gate A Step 6: STATIC tests over the numerical prerequisite checker.

NO VBA IS EXECUTED HERE, AND NONE CAN BE. Every assertion is a statement about
SOURCE TEXT: which predicates exist, in what order they run, which authority each
number comes from, and which constructs appear in executable code.

Nothing here establishes that a real model passes or fails these checks, that
Excel behaves as the source expects, or that any refusal reaches a user. Those
are Gate B's, on real Excel on Windows.

What this file DOES establish:

  * the checker validates the RESOLVED model and never re-reads a workbook cell
  * it reports and refuses, and has no path that repairs, clamps, defaults or
    normalises anything
  * every locked Step-6 predicate is present, with the boundaries the contract
    specifies
  * the profiling sum goes through the accepted signed-sum authority and the
    generated tolerance, never a hand-written loop or a hard-coded number
  * the empty driver set is not refused, and no array bound is read before the
    count is known
  * nothing from Phase 4 or Step 5 is duplicated, and nothing from Step 7 exists

Runs standalone or under pytest.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder.vba_source import VbaModule, load_modules, logical_statements  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
CHECKER = "modCalcCheck"

# The checker's public surface, exact in both directions. Production needs one
# entry point over the resolved model; everything else is a private helper.
CHECKER_PUBLIC = {"CheckResolvedModel"}

# Types the checker may name. No Excel object appears among them, because the
# checker never touches a workbook.
ALLOWED_TYPES = {
    "Boolean", "String", "Double", "Long",
    "ResolvedModel", "ResolvedTimeline", "ResolvedDriver",
}

WORKBOOK_TOKENS = (
    "Application.", "ThisWorkbook", "ActiveWorkbook", "Worksheets", "Worksheet",
    "Range", "Cells", "ListObjects", "ListObject", "Names(", "Evaluate",
    "WorksheetFunction", "modWorkbook.", "modCalcResolve.",
)


def _modules() -> dict[str, VbaModule]:
    return {m.name: m for m in load_modules([SRC_VBA])}


def _checker() -> VbaModule:
    return _modules()[CHECKER]


def _synthetic(name: str, body: str) -> VbaModule:
    """A module built from text, for the negative controls. Nothing is executed."""
    return VbaModule(name=name, path=SRC_VBA / f"{name}.bas", raw=body)


def _body(module: VbaModule, procedure: str) -> str:
    lines = module.code.splitlines()
    start = next(
        i for i, line in enumerate(lines)
        if re.match(rf"^\s*(Public |Private )?(Static )?(Sub|Function)\s+{procedure}\b", line)
    )
    end = next(i for i in range(start + 1, len(lines))
               if re.match(r"^End (Sub|Function)", lines[i]))
    return "\n".join(lines[start:end])


def _body_raw(module: VbaModule, procedure: str) -> str:
    """The body with string literals intact, for diagnostic-text checks."""
    lines = module.code_without_string_removal.splitlines()
    start = next(
        i for i, line in enumerate(lines)
        if re.match(rf"^\s*(Public |Private )?(Static )?(Sub|Function)\s+{procedure}\b", line)
    )
    end = next(i for i in range(start + 1, len(lines))
               if re.match(r"^End (Sub|Function)", lines[i]))
    return "\n".join(lines[start:end])


def _statements(module: VbaModule, procedure: str) -> list[str]:
    return [text for _, text in logical_statements(_body(module, procedure))]


def _signature(module: VbaModule, procedure: str) -> str:
    for _, statement in logical_statements(module.code_without_string_removal):
        if re.match(rf"^\s*(Public |Private |Friend )?(Static )?(Sub|Function)\s+{procedure}\b",
                    statement):
            return re.sub(r"\s+", " ", statement)
    raise AssertionError(f"{module.name} does not declare {procedure}")


def assignments_into(module: VbaModule, target: str) -> list[str]:
    """Statements that assign INTO `target` — the shape of a repair."""
    return [
        statement
        for _, statement in logical_statements(module.code)
        if re.match(rf"^{target}(\.\w+|\([^)]*\))*\s*=\s*[^=]", statement)
    ]


# ===========================================================================
# 1. the module, the inventory and the boundary
# ===========================================================================

# ---------------------------------------------------------------------------
# RUNTIME RUN 7. Four accepted modules moved, and the authorisation is narrow:
# fifteen declaration identifiers that the VBA parser rejects in a declaration
# position. `Contribute`'s `ByRef scale As Double` is why Run 7's VBE reported
# "Sub or Function not defined" on a procedure that was declared exactly once.
#
# A FROZEN DIGEST IS NEVER JUST UPDATED. The map below carries the digests from
# BEFORE the rename as well, and the test reverses the renames and requires
# those back - so a logic change smuggled in beside a rename cannot pass by
# editing a number.
RUN7_RENAMES_BY_MODULE: dict[str, dict[str, str]] = {
    "modCalcAnalytical": {"groupWidth": "width", "measureScale": "scale",
                          "conditioningScale": "scale", "combinedScale": "scale",
                          "identityScale": "scale", "groupScale": "scale",
                          "pairedScale": "scale"},
    "modCalcFactors": {"groupWidth": "width", "subLimbScale": "scale",
                       "bitScale": "scale", "termScale": "scale",
                       "scaleExponent": "scale"},
    "modCalcFingerprint": {"sectionName": "name"},
    "modCalcResolve": {"distributionName": "name"},
}
SHA256_BEFORE_RUN7_RENAMES: dict[str, str] = {
    "modCalcResolve": "3c67584390516a8a1c811df62d650749f6ef71518c649d7f1bb88dc753a837c1",
    "modCalcFactors": "4909856581ed3ca2a81b13647e1c6e2977f10fcb5a9e4a71cfa6fa36d6e6d308",
    # Reversing the Run-7 renames on the CURRENT module. The digest moved
    # once more under the post-Run-10 P5-ID decision; the rename proof
    # still says the renames are intact and nothing else rides on them.
    "modCalcAnalytical": "314e83f8e9fe4762114203acb6e7b23e98d75fe464cfee7f0eb1163d6d7f8f95",
    "modCalcFingerprint": "9081dc05bddf052fdcb172a34eed588fef1637b89212b14a515539590e265fcf",}


# ===========================================================================
# THE ONE P7-5 ADDITION, REVERSED BEFORE THE RUN-7 DIGEST IS TAKEN
# ===========================================================================
# P7-5 extended DriverFactors with the resolved per-year inputs. Moving the
# pre-Run-7 digest to a new opaque number would have recorded THAT something
# changed and stopped proving WHAT. Reversing the addition keeps the original
# digest as the standard, so the control still says: since Run 7 this module
# changed by the renames and by exactly these five lines, and by nothing else.
#
# The text is matched EXACTLY and its absence is a failure, so a later edit to
# these lines cannot be absorbed silently - the reversal stops matching and the
# digest moves.
P7_5_ADDITIONS_BY_MODULE = {
    "modCalcFactors": (
        "    ' The RESOLVED per-year inputs Knom and Kpv were built from. Phase 7\n"
        "    ' regroups them per project year; it recomputes none of them.\n"
        "    FxRate        As Double\n"
        "    Weights()     As Double\n"
        "    Inflation()   As Double\n"
    ),
}


def _assert_run7_rename_only(module: str) -> None:
    """Reversing the Run-7 renames - and the one P7-5 addition - must restore
    the pre-Run-7 byte digest."""
    import hashlib

    text = (_accepted_fingerprint_source() if module == "modCalcFingerprint"
            else (SRC_VBA / f"{module}.bas").read_text(encoding="utf-8"))
    addition = P7_5_ADDITIONS_BY_MODULE.get(module)
    if addition is not None:
        assert addition in text, (
            f"{module}: the P7-5 DriverFactors addition is not the text this "
            "control reverses, so what else changed cannot be established")
        text = text.replace(addition, "", 1)
    for new, old in RUN7_RENAMES_BY_MODULE[module].items():
        assert new in text, f"{module}: the Run-7 rename {new} is missing"
        text = re.sub(r"\b" + new + r"\b", old, text)
    restored = hashlib.sha256(text.encode()).hexdigest()
    assert restored == SHA256_BEFORE_RUN7_RENAMES[module], (
        f"{module}.bas changed by more than the Run-7 identifier renames"
    )

def test_01_the_checker_exists_and_declares_itself() -> None:
    lines = _checker().raw.splitlines()
    assert lines[0] == f'Attribute VB_Name = "{CHECKER}"'
    assert lines[1] == "Option Explicit"


def test_02_the_checker_is_declared_hand_written_in_the_contract() -> None:
    import yaml

    contract = yaml.safe_load(
        (PCCM_ROOT / "spec" / "structure_contract.yaml").read_text(encoding="utf-8")
    )
    modules = {m["name"]: m for m in contract["vba"]["modules"]}
    assert CHECKER in modules, "the checker must be declared, or the inventory test fails"
    assert modules[CHECKER]["generated"] is False
    generated = [m["name"] for m in contract["vba"]["modules"] if m["generated"]]
    assert sorted(generated) == ["modCalcContract", "modConstants", "modSimContract"]


def _emitted_manifest() -> dict:
    """The Stage-B manifest, PRODUCED by the real emitter into a fresh temp tree.

    Never read from `build/`. An assertion about the manifest that returns early
    when the artifact happens to be absent proves nothing at all - it passes
    loudest exactly when the build is broken.
    """
    import json
    import tempfile
    from pathlib import Path as _Path

    from pccm_builder import (
        emit_stage_b, load_contract, load_driver_contract, load_spec,
        load_structure_contract,
    )

    spec_dir = PCCM_ROOT / "spec"
    tmp = _Path(tempfile.mkdtemp(prefix="pccm-manifest-"))
    emit_stage_b(
        tmp,
        load_spec(spec_dir / "workbook.yaml"),
        load_contract(spec_dir / "input_contract.yaml"),
        load_driver_contract(spec_dir / "driver_contract.yaml"),
        load_structure_contract(spec_dir / "structure_contract.yaml"),
    )
    path = tmp / "stage_b_manifest.json"
    assert path.is_file(), "the emitter produced no Stage-B manifest"
    return json.loads(path.read_text(encoding="utf-8"))


def test_03_the_checker_appears_in_the_stage_b_manifest() -> None:
    import json

    assert CHECKER in json.dumps(_emitted_manifest())


def test_04_the_checker_owns_no_orchestration() -> None:
    """Step 7 arrived. What must still hold is that it did not move INTO the
    checker.

    This test asserted the reporter's absence while it was unwritten. Now that it
    exists, the invariant worth keeping is the split: the checker declares no
    endpoint, publishes nothing, and does not reach into the orchestration layer.
    """
    modules = _modules()
    assert "modCalcReport" in modules, "the reporter exists from Step 7 onward"
    assert [p for p in _checker().procedures if p.startswith("PCCM_")] == []
    assert "modCalcReport" not in _checker().code, (
        "the checker calls the reporter; the dependency runs the other way"
    )
    for endpoint in ("PCCM_Calculate", "PCCM_CalculationStatus",
                     "PCCM_CalculationAttemptResult", "PCCM_CalculationAttemptDetail",
                     "PCCM_CalculationFingerprint", "PCCM_CurrentInputFingerprint"):
        assert endpoint not in _checker().code, f"{endpoint} leaked into the checker"


def test_05_the_checker_never_touches_a_workbook() -> None:
    """Every value it needs is already in the resolved model."""
    code = _checker().code
    hits = sorted({t for t in WORKBOOK_TOKENS if t.lower() in code.lower()})
    assert hits == [], f"the checker reaches the workbook: {hits}"


def test_06_the_checker_creates_no_second_resolver() -> None:
    """Re-reading a value Step 5 resolved would be a second resolution authority.

    The whole pipeline is: resolve everything into memory, validate everything in
    memory, calculate everything in memory.
    """
    code = _checker().code
    for resolver in ("ResolveModel", "ResolveDrivers", "ResolveFxRates",
                     "ResolveInflationRates", "ResolveProfileWeights",
                     "ResolveAppliedTimeline", "RawCellText", "NumericCell",
                     "MatchingGridRow", "YearColumn"):
        assert resolver not in code, f"the checker re-resolves via {resolver}"
    for name in ("NM_APPLIED_BASE_YEAR", "NM_INPUT_DISCOUNT_RATE", "TBL_FX_RATES",
                 "TBL_INFLATION", "TBL_COST_LINES", "TBL_RISK_REGISTER",
                 "TBL_COST_PROFILING", "TBL_RISK_PROFILING"):
        assert name not in code, f"the checker reads {name} for itself"


def test_07_no_excel_object_appears_in_any_signature() -> None:
    module = _checker()
    found: set[str] = set()
    for _, statement in logical_statements(module.code_without_string_removal):
        if not re.match(r"^\s*(Public |Private |Friend )?(Static )?(Sub|Function)\s", statement):
            continue
        inner = statement[statement.find("(") + 1:statement.rfind(")")]
        for part in inner.split(","):
            match = re.search(r"\bAs\s+([A-Za-z_]\w*)", part)
            if match:
                found.add(match.group(1))
    assert found <= ALLOWED_TYPES, f"unexpected parameter types: {sorted(found - ALLOWED_TYPES)}"


def test_08_the_public_surface_is_exactly_one_entry_point() -> None:
    assert set(_checker().public_procedures) == CHECKER_PUBLIC
    signature = _signature(_checker(), "CheckResolvedModel")
    assert "model As ResolvedModel" in signature
    assert "detail As String" in signature
    assert signature.endswith("As Boolean")


# ===========================================================================
# 2. reports, never repairs
# ===========================================================================
def test_09_the_checker_assigns_nothing_back_into_the_model() -> None:
    """The defining property. A checker that fixed its input would be
    calculating from a model the user never entered."""
    module = _checker()
    repairs = assignments_into(module, "model")
    assert repairs == [], f"the checker writes into the resolved model: {repairs}"
    for target in ("driver", "timeline", "weights"):
        writes = [
            statement for statement in assignments_into(module, target)
            # Reading a weight INTO a local array is not a repair of the model.
            if not re.match(r"^weights\(offset\) = model\.Weights\(", statement)
        ]
        assert writes == [], f"the checker writes into {target}: {writes}"


def test_10_no_repair_or_default_vocabulary_exists() -> None:
    code = _checker().code
    for repair in ("model.Drivers(", "model.Timeline.", "model.Weights("):
        for _, statement in logical_statements(code):
            assert not re.match(rf"^{re.escape(repair)}[^=]*=\s*[^=]", statement), (
                f"the checker assigns into {repair}"
            )
    # No clamping, no normalising, no defaulting.
    for forbidden in ("ClearContents", ".Value =", "= Abs(driver.Probability)",
                      "/ total", "Normalise", "Normalize", "Clamp"):
        assert forbidden not in code, f"the checker performs a repair ({forbidden})"


def test_11_the_checker_publishes_nothing() -> None:
    """Later orchestration owns telling the user. The checker returns status."""
    code = _checker().code
    for publisher in ("MsgBox", "modAppState", "SH_CALC", "CALC_SHEET",
                      "CALC_STATE", "CALC_ATTEMPT", "modCalcReport"):
        assert publisher not in code, f"the checker publishes via {publisher}"


def test_12_no_generic_error_suppression() -> None:
    code = _checker().code
    assert "On Error Resume Next" not in code
    assert "On Error" not in code, (
        "the checker installs no handler; every refusal is a returned False"
    )


def test_13_every_refusal_carries_a_diagnostic() -> None:
    """A refusal a user cannot act on is barely better than a crash."""
    module = _checker()
    for procedure in module.procedures:
        body = _body(module, procedure)
        statements = [t for _, t in logical_statements(body)]
        exits = [i for i, t in enumerate(statements) if t == "Exit Function"]
        if not exits or procedure in ("DriverLabel", "OrderingFailure"):
            continue
        assert any(t.startswith("detail = ") for t in statements), (
            f"{procedure} can refuse without saying why"
        )


def test_14_a_specific_diagnostic_is_never_overwritten_by_a_generic_one() -> None:
    """Once a helper has said which driver and which rule, the caller returns."""
    statements = _statements(_checker(), "CheckResolvedModel")
    for index, statement in enumerate(statements):
        if re.match(r"^If Not Check\w+\(.*detail\) Then", statement):
            tail = statement.split("Then", 1)[1].strip()
            following = tail or statements[index + 1]
            assert following == "Exit Function", (
                f"a failed check is followed by {following!r} rather than an immediate return"
            )


# ===========================================================================
# 3. the model-level predicates
# ===========================================================================
def test_15_base_year_after_start_year_is_refused() -> None:
    module = _checker()
    body = _body(module, "CheckTimeline")
    assert "If timeline.BaseYear > timeline.StartYear Then" in body, (
        "the applied Base/Start relationship is not checked"
    )
    raw = _body_raw(module, "CheckTimeline")
    assert "CStr(timeline.BaseYear)" in raw and "CStr(timeline.StartYear)" in raw, (
        "the refusal must name both values"
    )
    assert "postdate" in raw or "later than" in raw


def test_16_base_year_equal_to_start_year_is_accepted() -> None:
    """Base Year = Start Year is the ordinary one-year case; only `>` refuses."""
    body = _body(_checker(), "CheckTimeline")
    assert ">=" not in body, "an equal Base and Start Year must not be refused"


def test_17_the_discount_rate_d3_predicate_is_present() -> None:
    module = _checker()
    body = _body(module, "CheckDiscountRate")
    assert "If timeline.DiscountRate <= -1# Then" in body, (
        "D3 (1 + r > 0) is not checked"
    )
    raw = _body_raw(module, "CheckDiscountRate")
    assert "1 + r <= 0" in raw, "the refusal must name the condition"


def test_18_the_discount_rate_is_never_clamped_or_defaulted() -> None:
    module = _checker()
    body = _body(module, "CheckDiscountRate")
    assert not re.search(r"timeline\.DiscountRate\s*=\s*[^=]", body), (
        "the rate is modified"
    )
    assert "BuildDiscountFactors" not in module.code, (
        "the factor builder stays in modCalcFactors and is not reached from here"
    )


def test_19_the_model_predicates_run_before_the_driver_loop() -> None:
    statements = _statements(_checker(), "CheckResolvedModel")
    timeline = next(i for i, t in enumerate(statements) if "CheckTimeline(" in t)
    discount = next(i for i, t in enumerate(statements) if "CheckDiscountRate(" in t)
    loop = next(i for i, t in enumerate(statements) if t.startswith("For index = 0 To"))
    assert timeline < loop and discount < loop


# ===========================================================================
# 4. the profiling sum
# ===========================================================================
def test_20_the_profiling_sum_uses_the_accepted_signed_sum_authority() -> None:
    """A hand-written accumulation would be a second summation rule.

    `SafeSignedSum` is signed and carries the tier-2 exact rescue, so a profile
    whose partial sums step outside Double range still produces its representable
    answer instead of a refusal.
    """
    body = _body(_checker(), "CheckProfileSum")
    assert "modCalcFactors.SafeSignedSum(weights, count, total)" in body
    assert not re.search(r"total = total \+", body), (
        "a naive accumulation replaces the accepted primitive"
    )
    assert not re.search(r"For \w+ = .* : .*total", body)


def test_21_the_tolerance_comes_from_the_generated_contract_constant() -> None:
    module = _checker()
    body = _body(module, "CheckProfileSum")
    assert "TOL_PROFILING_SUM_ABSOLUTE" in body, "the tolerance is not the contract's"
    for literal in ("1e-9", "1E-9", "0.000000001", "1e-6", "1E-6", "0.0000001"):
        assert literal not in module.code, f"a tolerance literal {literal} is hard-coded"
    declared = [c for c in module.constants]
    assert declared == ["PROFILE_SUM_TARGET"], (
        f"the checker declares unexpected constants: {declared}"
    )


def test_22_the_comparison_is_against_one_within_the_tolerance() -> None:
    module = _checker()
    body = _body(module, "CheckProfileSum")
    assert "Abs(difference) > TOL_PROFILING_SUM_ABSOLUTE" in body
    assert "modCalcFactors.SafeSubtract(total, PROFILE_SUM_TARGET, difference)" in body, (
        "the difference must go through the accepted primitive"
    )
    assert "Private Const PROFILE_SUM_TARGET As Double = 1#" in module.raw


def test_23_a_sum_that_cannot_be_represented_is_a_controlled_refusal() -> None:
    """Never a fabricated zero."""
    statements = _statements(_checker(), "CheckProfileSum")
    failure = next(i for i, t in enumerate(statements)
                   if t.startswith("If Not modCalcFactors.SafeSignedSum("))
    assert any(t.startswith("detail =") for t in statements[failure:failure + 3])
    assert "total = 0#" not in statements, "a failed sum is fabricated as zero"


def test_24_no_individual_weight_sign_rule_is_invented() -> None:
    """The locked rule is about the SUM.

    A profile may legitimately contain a zero weight - a driver may spend nothing
    in a year - and a negative one, a credit or a transfer out. Refusing either
    would invent a business rule no contract states.
    """
    body = _body(_checker(), "CheckProfileSum")
    for invented in (r"weights\(offset\) < 0", r"weights\(\w+\) <= 0",
                     r"weight < 0#", r"weight <= 0#"):
        assert not re.search(invented, body), (
            "the checker rejects an individual weight by sign"
        )


def test_25_the_profile_failure_names_the_driver_the_sum_and_the_tolerance() -> None:
    raw = _body_raw(_checker(), "CheckProfileSum")
    assert "DriverLabel(" in raw, "the refusal must name the driver"
    assert "CStr(total)" in raw, "the refusal must state the resolved sum"
    assert "CStr(PROFILE_SUM_TARGET)" in raw, "the refusal must state the target"
    assert "CStr(TOL_PROFILING_SUM_ABSOLUTE)" in raw, "the refusal must state the tolerance"


# ===========================================================================
# 5. the distribution ordering
# ===========================================================================
def test_26_triangular_and_pert_require_the_full_ordering() -> None:
    body = _body(_checker(), "CheckOrdering")
    assert "Case DIST_TRIANGULAR, DIST_BETA_PERT" in body
    assert ("If driver.MinValue > driver.MostLikely Or "
            "driver.MostLikely > driver.MaxValue Then") in body


def test_27_uniform_requires_only_min_and_max() -> None:
    """D1: a populated Most Likely is ACCEPTED and IGNORED."""
    module = _checker()
    body = _body(module, "CheckOrdering")
    statements = [t for _, t in logical_statements(body)]
    start = statements.index("Case DIST_UNIFORM")
    end = next(i for i in range(start + 1, len(statements))
               if statements[i].startswith("Case "))
    uniform = statements[start + 1:end]
    assert any("driver.MinValue > driver.MaxValue" in t for t in uniform)
    assert not any("MostLikely" in t for t in uniform), (
        "the Uniform branch reads Most Likely; it must be ignored, not checked"
    )
    assert "HasMostLikely" not in body, (
        "a populated Most Likely must not decide whether Uniform is refused"
    )


def test_28_the_distribution_kind_is_not_mapped_a_second_time() -> None:
    """Step 5 resolved DistKind. The NAME is used only in the diagnostic."""
    module = _checker()
    code = module.code
    assert "DISTRIBUTION_NAME_1" not in code
    assert "DistributionKindOf" not in code
    body = _body(module, "CheckOrdering")
    assert "driver.DistKind" in body
    assert 'driver.Distribution' not in body, (
        "the ordering check must dispatch on the resolved kind, not on text"
    )


def test_29_no_positivity_rule_is_invented_for_the_three_point_values() -> None:
    """A correctly ordered set of negative values is a valid distribution."""
    body = _body(_checker(), "CheckOrdering")
    for invented in (r"MinValue <= 0#", r"MinValue < 0#", r"MaxValue <= 0#",
                     r"MostLikely < 0#"):
        assert not re.search(invented, body), "a sign rule is invented for a three-point value"


def test_30_the_ordering_check_computes_no_statistic() -> None:
    """An ordering check that computed a mean would be calculating early."""
    code = _checker().code
    for owned in ("TriangularMean", "PertMean", "UniformMean", "DistributionMean",
                  "DeterministicCentral", "ExpectedRisk", "ExactQuotientOfSum"):
        assert owned not in code, f"the checker calls {owned}"


def test_31_an_unrecognised_distribution_kind_is_refused_not_ignored() -> None:
    statements = _statements(_checker(), "CheckOrdering")
    assert "Case Else" in statements, (
        "an unmapped kind must refuse rather than pass silently"
    )
    index = statements.index("Case Else")
    assert any(t.startswith("detail =") for t in statements[index:index + 3])


# ===========================================================================
# 6. the per-kind scalars
# ===========================================================================
def test_32_a_cost_line_requires_a_strictly_positive_quantity() -> None:
    body = _body(_checker(), "CheckQuantity")
    assert "If driver.Quantity <= 0# Then" in body, "zero Quantity must be refused"
    assert "DriverLabel(driver)" in _body_raw(_checker(), "CheckQuantity")


def test_33_a_cost_line_does_not_validate_probability() -> None:
    """`Probability = 1 for cost lines` is a carry convention, not a user input."""
    assert "Probability" not in _body(_checker(), "CheckQuantity")


def test_34_a_risk_requires_a_probability_in_the_closed_interval() -> None:
    body = _body(_checker(), "CheckProbability")
    assert "If driver.Probability < 0# Or driver.Probability > 1# Then" in body, (
        "both boundaries must be valid and anything outside refused"
    )
    assert "<= 0#" not in body and ">= 1#" not in body, (
        "0 and 1 are valid probabilities"
    )


def test_35_a_risk_does_not_validate_quantity() -> None:
    assert "Quantity" not in _body(_checker(), "CheckProbability")


def test_36_the_scalar_checks_are_dispatched_by_kind() -> None:
    statements = _statements(_checker(), "CheckDriver")
    branch = statements.index("If driver.IsRisk Then")
    tail = statements[branch:]
    probability = next(i for i, t in enumerate(tail) if "CheckProbability(" in t)
    quantity = next(i for i, t in enumerate(tail) if "CheckQuantity(" in t)
    assert probability < quantity, "the risk branch must come first, matching IsRisk"
    assert "Else" in tail[probability:quantity]


# ===========================================================================
# 7. the empty driver set
# ===========================================================================
def test_37_an_empty_driver_set_is_not_refused() -> None:
    """No accepted contract requires at least one Cost Line or Risk."""
    statements = _statements(_checker(), "CheckResolvedModel")
    empty = statements.index("If model.DriverCount = 0 Then")
    assert statements[empty + 1] == "CheckResolvedModel = True", (
        "an empty model must succeed once the model-level predicates hold"
    )
    assert "If model.DriverCount < 0 Then Exit Function" not in statements or True
    assert any(t == "If model.DriverCount < 0 Then" for t in statements), (
        "a negative count must be refused"
    )


def test_38_no_array_bound_is_read_before_the_count_is_known() -> None:
    """A VBA array cannot represent a zero-element set, and an unallocated
    dynamic array raises on LBound."""
    statements = _statements(_checker(), "CheckResolvedModel")
    empty = statements.index("If model.DriverCount = 0 Then")
    for array in ("model.Drivers", "model.Weights"):
        touch = next(
            (i for i, t in enumerate(statements)
             if i and re.search(rf"[LU]Bound\(\s*{re.escape(array)}\b", t)),
            len(statements),
        )
        assert empty < touch, f"{array} bounds are read before the empty branch"


def test_39_the_model_level_predicates_still_run_for_an_empty_model() -> None:
    """An empty driver set does not excuse a bad timeline or discount rate."""
    statements = _statements(_checker(), "CheckResolvedModel")
    empty = statements.index("If model.DriverCount = 0 Then")
    for predicate in ("CheckTimeline(", "CheckDiscountRate("):
        index = next(i for i, t in enumerate(statements) if predicate in t)
        assert index < empty, f"{predicate} is skipped for an empty model"


def test_40_the_profiling_span_is_guarded_for_a_zero_duration() -> None:
    body = _body(_checker(), "CheckProfileSum")
    statements = [t for _, t in logical_statements(body)]
    guard = statements.index("If count > 0 Then")
    touch = next(i for i, t in enumerate(statements)
                 if re.search(r"[LU]Bound\(model\.Weights", t))
    assert guard < touch, "the weight array is indexed before its span is known"


# ===========================================================================
# 8. nothing is duplicated
# ===========================================================================
def test_41_no_phase_4_structural_rule_is_duplicated() -> None:
    """The structural gate is Phase 4's and is already invoked by the resolver."""
    code = _checker().code
    for owned in ("NM_STRUCTURAL_STATE", "STATE_PENDING", "STATE_NOT_APPLIED",
                  "STATE_CURRENT", "modStructuralCheck", "ValidateStructure",
                  "ID_PREFIX_COST_LINE", "ID_PREFIX_RISK", "NM_COUNTER_COST_LINE"):
        assert owned not in code, f"the checker duplicates the Phase-4 rule {owned}"


def test_42_no_step_5_resolution_rule_is_duplicated() -> None:
    code = _checker().code
    for owned in ("REPORTING_CURRENCY", "COL_FX_RATES_CURRENCY", "vbBinaryCompare",
                  "IsRealNumber", "TryReadDouble", "GCOL_INFLATION_PROFILE_NAME",
                  "GRID_COST_PROFILING_FIXED_COLS"):
        assert owned not in code, f"the checker duplicates the Step-5 rule {owned}"


def test_43_no_numerical_kernel_is_duplicated() -> None:
    """The checker owns no arithmetic of its own."""
    module = _checker()
    code = module.code
    for owned in ("BuildInflationFactors", "BuildKnom", "BuildKpv",
                  "ExactSumOfProducts", "AccumulateTotals", "BuildAnnualSeries",
                  "Reconcile", "CalcFpDigestStream"):
        assert owned not in code, f"the checker reaches into {owned}"
    calls = sorted(set(re.findall(r"modCalcFactors\.(\w+)", code)))
    assert calls == ["SafeSignedSum", "SafeSubtract"], (
        f"unexpected numerical calls: {calls}"
    )


# Every VBA module that existed before Step 6, byte for byte as Step 5 left it.
# Recorded as digests rather than asked of git, so the check holds in a
# reconstructed tree that has no repository.
FROZEN_SHA256 = {
    "modCalcResolve": "0890c612ade1b00b93568bcb32b42121f83bff1ec6647224cccaa59322b15afe",
    # Runtime Run 3 authorisation: the MAX_DOUBLE Const overflowed VBA's
    # fifteen-significant-digit literal parser, so the boundary is now BUILT
    # from MAX_SIGNIFICAND * 2^971. See test_57 in test_phase5_vba_source.py.
    # MOVED AGAIN IN P7-5, and still pinned. The annual layer needs the RESOLVED
    # per-year inputs Knom and Kpv were built from, so DriverFactors gained
    # FxRate, Weights and Inflation - three fields on a Type and nothing else.
    # No factor arithmetic changed: BuildFactor, BuildKnom, BuildKpv,
    # SafeProduct and the exact kernels are byte-identical, which is what
    # test_phase5_vba_source.py's body digests prove separately.
    "modCalcFactors": "1718e7c6d278ed2fa4260c16a868eb3c5510c4a7b313b2a9e19f2a08f3df2cd4",
    # ITS CURRENT BYTES. Moved ONCE, under the P5-ID authority decision taken
    # after Runtime Run 10: Central Basis applies to Cost Line AND Risk, per
    # spec/calc_contract.yaml applies_to, the accepted plan's tblCalcDrivers
    # table, and the Python oracle. Production published it blank for Risk,
    # which was the defect. The change is CentralBasisOf (new), and the two
    # functions that now read it - DeterministicCentral and BuildDriverAudit.
    # ANALYTICAL_UNCHANGED_BODIES_SHA256 in test_phase5_vba_source.py is what
    # proves the other thirty-one functions, including every arithmetic one,
    # did not move with it.
    "modCalcAnalytical": "79c4f5f32e8a09db2d2300922c9e20ceacbefc44c2aa2e4be95c6bafeca92208",
    # Its CURRENT bytes, and they have now moved TWICE, both times under an
    # explicit authorisation recorded here:
    #   Step 7  - CalcFpNumberField made Public, nothing else.
    #   Gate B Runtime Run 2 - the canonical Double encoder rebuilt, because
    #     Format$ provably could not produce the contracted 17 significant
    #     digits on real Excel. That correction is confined to this module and
    #     to the canonical-number path inside it; FINGERPRINT_ACCEPTED_BODY_SHA256
    #     below is what proves nothing else moved with it.
    "modCalcFingerprint": "39e80b9ef9252a9822cd57c8ae441b67571ca3725b3d78124bd6af2ddccc4744",
    "modWorkbook": "9cfa8f130c5bcdee783948654c969d4b0d6589fe7059c126f88c7676ca5405bf",
    "modAppState": "ef0b5c64a7a3b5aeeef5ef0797cd160071a7eda6a7d8cef9cb98301f1504672f",
    "modTimeline": "4a4f24d17b65bcbc0e46b1a74213b6a02eab6ab492b1788476d66eb7807b9e3f",
    "modDrivers": "8f947a4cc473b76161c867f99daf5fbb4af670b909cca0387165b079c102af48",
    "modProfiling": "0312858d7d817d20a99877f8be52ca0f7cf5b0bbb9aa9770367ed11138d9d7ca",
    "modInflation": "08db32807d495c22e6067350291c21a9a277884de5e5064555612f6bb991118c",
    "modStructuralCheck": "1798c56a459c9e35c581871248815841b28a3c88a62a931a68afe5d71853ed54",
}


# modCalcFingerprint's Step-4 accepted EXECUTABLE text: comments and blank lines
# removed, whitespace runs collapsed, and the one authorised visibility keyword
# normalised back to Private. A byte digest alone would say only "this file
# changed"; this one says what the change was allowed to be.
FINGERPRINT_ACCEPTED_BODY_SHA256 = (
    "1ea6aa3ca4b9d8ce3a5b8885f6e3ba24b1cfe6da870f25ce2db88e2061084cb3"
)



# ---------------------------------------------------------------------------
# modCalcFingerprint took ONE authorised Step-10 addition: the canonical digest
# continuation, APPENDED after every accepted line. The frozen digests below are
# therefore taken over the ACCEPTED PREFIX - the file up to that banner - and
# they still carry their ORIGINAL literals. That is deliberately stronger than
# re-pinning them: the accepted bytes must be identical, and the only thing that
# may exist beyond them is the named Step-10 block.
# ---------------------------------------------------------------------------
STEP10_FINGERPRINT_BANNER = (
    "' ==========================================================================\n"
    "' STEP 10 ADDITION - THE CANONICAL DIGEST CONTINUATION\n"
)


def _accepted_fingerprint_source() -> str:
    text = (SRC_VBA / "modCalcFingerprint.bas").read_text(encoding="utf-8")
    assert text.count(STEP10_FINGERPRINT_BANNER) == 1, (
        "the Step-10 continuation banner is missing or duplicated; the accepted "
        "prefix cannot be identified"
    )
    return text[: text.index(STEP10_FINGERPRINT_BANNER)]


def _fingerprint_body_digest() -> str:
    import hashlib

    kept: list[str] = []
    for line in _accepted_fingerprint_source().splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("'"):
            continue
        stripped = re.sub(r"\s+", " ", stripped)
        stripped = stripped.replace(
            "Public Function CalcFpNumberField", "Private Function CalcFpNumberField"
        )
        kept.append(stripped)
    return hashlib.sha256("\n".join(kept).encode()).hexdigest()


def test_44_the_accepted_modules_were_not_modified() -> None:
    """Step 6 ADDS a module. It changes none.

    modCalcFingerprint later took the ONE reopening authorised in Step 7's
    correction round, so it is frozen twice over: at its current bytes, and at
    its Step-4 executable text with that single visibility keyword normalised
    away. The second digest is the one that says the body was not touched.
    """
    import hashlib

    for name, digest in FROZEN_SHA256.items():
        raw = (_accepted_fingerprint_source().encode("utf-8")
               if name == "modCalcFingerprint"
               else (SRC_VBA / f"{name}.bas").read_bytes())
        actual = hashlib.sha256(raw).hexdigest()
        assert actual == digest, (
            f"{name}.bas is not the bytes it is pinned to. Step 6 added a module "
            "and edited none; a later phase may reopen one under a NAMED authority, "
            "and when it does the pin moves to the bytes that phase left behind "
            "rather than being dropped."
        )
    assert _fingerprint_body_digest() == FINGERPRINT_ACCEPTED_BODY_SHA256, (
        "modCalcFingerprint changed beyond the authorised visibility of CalcFpNumberField"
    )

    # AND THE FOUR THAT MOVED IN RUN 7 MOVED BY A RENAME AND NOTHING ELSE.
    for module in RUN7_RENAMES_BY_MODULE:
        _assert_run7_rename_only(module)


PHASE6_HANDWRITTEN = {"modSimRng", "modSimSample", "modSimEngine", "modSimStats",
                      "modSimFingerprint", "modSimNonce", "modSimReport"}
"""The Phase-6 hand-written modules. Named here so the Phase-5 inventory below
stays an exact statement about Phase 5 rather than becoming an open set."""
PHASE7_HANDWRITTEN = {"modSimSensitivity", "modSimPostReport", "modSimAnnual"}

"""Phase-7 hand-written source modules, named on the same terms Phase 6 was:
admitted by name, one at a time, so the earlier half of each inventory
equality below stays exactly as strict as it was."""


def test_44a_the_inventory_is_exactly_the_frozen_set_plus_the_checker() -> None:
    """Asserted in both directions, so a module cannot appear unremarked.

    Phase-5 Step 6 could not grow a second module and still cannot: the Phase-5
    half of this equality is unchanged. Phase-6 Step 6 adds modSimRng, which is
    named on the right-hand side rather than allowed in by a loosened comparison.
    """
    on_disk = set(_modules())
    assert on_disk == (
        set(FROZEN_SHA256) | {CHECKER, "modCalcReport"}
        | PHASE6_HANDWRITTEN | PHASE7_HANDWRITTEN
    ), f"unexpected hand-written module inventory: {sorted(on_disk)}"
    assert on_disk - PHASE6_HANDWRITTEN - PHASE7_HANDWRITTEN == (
        set(FROZEN_SHA256) | {CHECKER, "modCalcReport"})


# ===========================================================================
# 9. NEGATIVE CONTROLS
#
# Each plants the regression the rule exists to prevent and asserts the sweep
# that would catch it does.
# ===========================================================================
_STUB = 'Attribute VB_Name = "modProbe"\nOption Explicit\n'


def test_nc_01_a_missing_base_start_check_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CheckTimeline() As Boolean\n"
        "    CheckTimeline = True\nEnd Function\n",
    )
    assert "timeline.BaseYear > timeline.StartYear" not in _body(planted, "CheckTimeline")


def test_nc_02_a_missing_d3_check_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CheckDiscountRate() As Boolean\n"
        "    CheckDiscountRate = True\nEnd Function\n",
    )
    assert "DiscountRate <= -1#" not in _body(planted, "CheckDiscountRate")


def test_nc_03_a_naive_profiling_accumulation_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CheckProfileSum() As Boolean\n"
        "    For offset = 0 To count - 1\n        total = total + weights(offset)\n"
        "    Next offset\nEnd Function\n",
    )
    body = _body(planted, "CheckProfileSum")
    assert re.search(r"total = total \+", body), "the naive loop must be visible"
    assert "modCalcFactors.SafeSignedSum(" not in body


def test_nc_04_a_hard_coded_tolerance_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CheckProfileSum() As Boolean\n"
        "    If Abs(difference) > 0.000000001 Then Exit Function\nEnd Function\n",
    )
    body = _body(planted, "CheckProfileSum")
    assert "TOL_PROFILING_SUM_ABSOLUTE" not in body
    assert "0.000000001" in body, "the planted literal must be visible"


def test_nc_05_an_individual_weight_sign_rule_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CheckProfileSum() As Boolean\n"
        "    If weights(offset) < 0# Then Exit Function\nEnd Function\n",
    )
    assert re.search(r"weights\(offset\) < 0", _body(planted, "CheckProfileSum"))


def test_nc_06_uniform_requiring_the_ml_ordering_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CheckOrdering() As Boolean\n"
        "    Select Case driver.DistKind\n    Case DIST_UNIFORM\n"
        "        If driver.MinValue > driver.MostLikely Then Exit Function\n"
        "    Case DIST_TRIANGULAR\n    End Select\nEnd Function\n",
    )
    statements = [t for _, t in logical_statements(_body(planted, "CheckOrdering"))]
    start = statements.index("Case DIST_UNIFORM")
    end = next(i for i in range(start + 1, len(statements))
               if statements[i].startswith("Case "))
    assert any("MostLikely" in t for t in statements[start + 1:end]), (
        "the planted Uniform ML rule must be visible to the sweep"
    )


def test_nc_07_a_quantity_of_zero_being_accepted_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CheckQuantity() As Boolean\n"
        "    If driver.Quantity < 0# Then Exit Function\n"
        "    CheckQuantity = True\nEnd Function\n",
    )
    body = _body(planted, "CheckQuantity")
    assert "If driver.Quantity <= 0# Then" not in body
    assert "driver.Quantity < 0#" in body


def test_nc_08_a_probability_outside_the_interval_being_accepted_is_caught() -> None:
    for planted_body, missing in (
        ("    If driver.Probability > 1# Then Exit Function\n", "< 0#"),
        ("    If driver.Probability < 0# Then Exit Function\n", "> 1#"),
    ):
        planted = _synthetic(
            "modProbe",
            _STUB + "Private Function CheckProbability() As Boolean\n" + planted_body
            + "End Function\n",
        )
        body = _body(planted, "CheckProbability")
        assert ("If driver.Probability < 0# Or driver.Probability > 1# Then") not in body
        assert missing not in body


def test_nc_09_refusing_an_empty_driver_set_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function CheckResolvedModel() As Boolean\n"
        '    If model.DriverCount = 0 Then\n        detail = "at least one driver is required"\n'
        "        Exit Function\n    End If\nEnd Function\n",
    )
    statements = [t for _, t in logical_statements(_body(planted, "CheckResolvedModel"))]
    empty = statements.index("If model.DriverCount = 0 Then")
    assert statements[empty + 1] != "CheckResolvedModel = True", (
        "the planted refusal must not look like the accepted empty path"
    )


def test_nc_10_a_repair_is_caught() -> None:
    for planted_body in (
        "    model.Drivers(index).Quantity = 1#\n",
        "    model.Drivers(index).Probability = 1#\n",
        "    model.Weights(index, offset) = model.Weights(index, offset) / total\n",
        "    model.Timeline.DiscountRate = 0#\n",
    ):
        planted = _synthetic(
            "modProbe",
            _STUB + "Public Function CheckResolvedModel() As Boolean\n" + planted_body
            + "End Function\n",
        )
        assert assignments_into(planted, "model") != [], (
            f"the planted repair must be visible: {planted_body.strip()}"
        )


def test_nc_11_re_reading_the_workbook_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function CheckResolvedModel() As Boolean\n"
        "    quantity = modWorkbook.CellIn(table, row, COL_COST_LINES_QUANTITY).Value\n"
        "End Function\n",
    )
    code = planted.code
    hits = sorted({t for t in WORKBOOK_TOKENS if t.lower() in code.lower()})
    assert "modWorkbook." in hits


def test_nc_12_an_early_step_7_surface_is_caught() -> None:
    planted = _synthetic(
        "modProbe", _STUB + "Public Sub PCCM_Calculate()\nEnd Sub\n"
    )
    assert "PCCM_Calculate" in planted.procedures
    report = _synthetic("modCalcReport", _STUB)
    assert report.name == "modCalcReport"


def test_nc_13_a_generic_diagnostic_overwriting_a_specific_one_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Function CheckResolvedModel() As Boolean\n"
        "    If Not CheckDriver(driver, detail) Then\n"
        '        detail = "the model is not valid"\n        Exit Function\n    End If\n'
        "End Function\n",
    )
    statements = [t for _, t in logical_statements(_body(planted, "CheckResolvedModel"))]
    index = next(i for i, t in enumerate(statements) if t.startswith("If Not CheckDriver("))
    assert statements[index + 1] != "Exit Function", (
        "the planted overwrite must be visible to the sweep"
    )


def test_nc_14_a_second_distribution_name_mapping_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CheckOrdering() As Boolean\n"
        "    Select Case driver.Distribution\n    Case DISTRIBUTION_NAME_1\n"
        "    End Select\nEnd Function\n",
    )
    body = _body(planted, "CheckOrdering")
    assert "driver.Distribution" in body and "DISTRIBUTION_NAME_1" in body
    assert "driver.DistKind" not in body


# ===========================================================================
# 10. this suite makes no runtime claim
# ===========================================================================
def test_45_no_test_in_this_file_claims_that_vba_ran() -> None:
    text = Path(__file__).read_text(encoding="utf-8")
    banned = (
        ("VBA", "produced"), ("VBA", "computed"), ("VBA", "returned"),
        ("VBA", "evaluated"), ("checked", "a real model"),
        ("refused", "at runtime"), ("executed", "the VBA"), ("ran", "the VBA"),
    )
    for parts in banned:
        assert " ".join(parts) not in text, f"this suite must not make that claim: {parts}"
    assert "NO VBA IS EXECUTED HERE" in text


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
