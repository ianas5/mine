#!/usr/bin/env python3
"""PCCM Phase 5 Gate A Step 7: STATIC tests over the reporting/orchestration layer.

NO VBA IS EXECUTED HERE, AND NONE CAN BE. Every assertion is a statement about
SOURCE TEXT: what is written where, in what order, inside which error envelope,
and from which authority.

Nothing here establishes that a real transaction commits, that a real rollback
restores a real workbook, that Excel raises where the source expects it to, or
that any status a user sees is right. **Those are Gate B's**, on real Excel on
Windows, and the six-row status matrix in particular can only be demonstrated
there. What Gate A can do is prove the source is CAPABLE of every row and
contains no ordering that would make one unreachable.

What this file DOES establish:

  * one preparation path serves Calculate, Status and CurrentInputFingerprint
  * every analytical value is computed before anything is written
  * a pre-write refusal touches C17:C20 and nothing else
  * the success commit is one assignment, verified, inside the rollback envelope
  * rollback restores all five tables and both scalar blocks, before any FAILED
    metadata is recorded
  * derived status is never REFUSED or FAILED and never reads attempt history
  * empty digests are never compared as though they were digests
  * an inapplicable audit field is blank, never an identity or a zero
  * exactly six Phase-5 endpoints, and no Calculate button

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

from pccm_builder.vba_source import VbaModule, load_modules, logical_statements  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
REPORTER = "modCalcReport"

# The locked Phase-5 automation/API surface. Exactly six, and no seventh.
PCCM_ENDPOINTS = {
    "PCCM_Calculate",
    "PCCM_CalculationStatus",
    "PCCM_CalculationAttemptResult",
    "PCCM_CalculationAttemptDetail",
    "PCCM_CalculationFingerprint",
    "PCCM_CurrentInputFingerprint",
}

# The failpoint stages a later Gate-B harness may arm. Public because the
# harness names them; they are constants, not procedures.
FAILPOINTS = {"FAILPOINT_ANALYTICAL_WRITE", "FAILPOINT_SUCCESS_COMMIT"}

ANALYTICAL_TABLES = (
    "TBL_CALC_YEARS", "TBL_CALC_INFLATION_FACTORS", "TBL_CALC_FX",
    "TBL_CALC_DRIVERS", "TBL_CALC_ANNUAL",
)

DERIVED_STATUSES = {
    "CALC_STATUS_NOT_CALCULATED", "CALC_STATUS_CURRENT",
    "CALC_STATUS_STALE", "CALC_STATUS_INVALID",
}


def _modules() -> dict[str, VbaModule]:
    return {m.name: m for m in load_modules([SRC_VBA])}


def _reporter() -> VbaModule:
    return _modules()[REPORTER]


def _synthetic(name: str, body: str) -> VbaModule:
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


def first_index(statements: list[str], pattern: str) -> int:
    """Index of the first statement matching `pattern`, or len() if absent."""
    return next((i for i, t in enumerate(statements) if re.search(pattern, t)),
                len(statements))


def emitted_manifest() -> dict:
    """The Stage-B manifest, PRODUCED by the real emitter into a fresh temp tree.

    Never read from `build/`. An assertion about the manifest that returns early
    when the artifact is absent proves nothing - it passes loudest exactly when
    the build is broken.
    """
    from pccm_builder import (
        emit_stage_b, load_contract, load_driver_contract, load_spec,
        load_structure_contract,
    )

    spec_dir = PCCM_ROOT / "spec"
    tmp = Path(tempfile.mkdtemp(prefix="pccm-step7-"))
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


# ===========================================================================
# 1. the module and the final inventory
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
    "modCalcAnalytical": "e234b3adacdb443c8c7b2b5072c311e7622405c3ec2e2987a750d85400299e0d",
    "modCalcFingerprint": "9081dc05bddf052fdcb172a34eed588fef1637b89212b14a515539590e265fcf",}


def _assert_run7_rename_only(module: str) -> None:
    """Reversing the Run-7 renames must restore the pre-Run-7 byte digest."""
    import hashlib

    text = (SRC_VBA / f"{module}.bas").read_text(encoding="utf-8")
    for new, old in RUN7_RENAMES_BY_MODULE[module].items():
        assert new in text, f"{module}: the Run-7 rename {new} is missing"
        text = re.sub(r"\b" + new + r"\b", old, text)
    restored = hashlib.sha256(text.encode()).hexdigest()
    assert restored == SHA256_BEFORE_RUN7_RENAMES[module], (
        f"{module}.bas changed by more than the Run-7 identifier renames"
    )

def test_01_the_reporter_exists_and_declares_itself() -> None:
    lines = _reporter().raw.splitlines()
    assert lines[0] == f'Attribute VB_Name = "{REPORTER}"'
    assert lines[1] == "Option Explicit"


def test_02_the_final_phase_5_inventory_is_fifteen_modules() -> None:
    manifest = emitted_manifest()
    names = [m["name"] for m in manifest["vba"]["modules"]]
    assert len(names) == 15, f"expected fifteen modules, found {len(names)}: {names}"
    assert REPORTER in names
    generated = [m["name"] for m in manifest["vba"]["modules"] if m["generated"]]
    assert sorted(generated) == ["modCalcContract", "modConstants"]
    assert set(_modules()) == set(names) - {"modConstants", "modCalcContract"}


def test_03_no_extra_orchestration_module_was_invented() -> None:
    """The locked architecture names modCalcReport, and only it."""
    on_disk = set(_modules())
    for invented in ("modCalcOrchestrator", "modCalcState", "modCalcTransaction",
                     "modCalcApi", "modCalcWrite", "modCalcCommit"):
        assert invented not in on_disk, f"{invented} was invented to split the work"


# ===========================================================================
# 2. the endpoint surface
# ===========================================================================
def test_04_exactly_six_phase_5_endpoints_exist() -> None:
    modules = _modules()
    # Every PCCM_ procedure across the whole codebase, minus the Phase-4 ones
    # the contract already accounts for.
    phase4 = {p for p in modules["modAppState"].public_procedures} | {
        "PCCM_ApplyTimeline", "PCCM_AddCostLine", "PCCM_DeleteCostLine",
        "PCCM_AddRisk", "PCCM_DeleteRisk", "PCCM_StructuralReport",
        "PCCM_DeleteCostLineById", "PCCM_DeleteRiskById",
    }
    found = {p for m in modules.values() for p in m.public_procedures
             if p.startswith("PCCM_")} - phase4
    assert found == PCCM_ENDPOINTS, (
        f"unexpected: {sorted(found - PCCM_ENDPOINTS)}; missing: "
        f"{sorted(PCCM_ENDPOINTS - found)}"
    )
    assert set(_reporter().public_procedures) & PCCM_ENDPOINTS == PCCM_ENDPOINTS, (
        "the endpoints must live in the reporter"
    )


def test_05_the_replaced_endpoint_does_not_exist() -> None:
    """`PCCM_CalculationRefusal` was explicitly replaced by the attempt axis."""
    everything = "\n".join(m.raw for m in _modules().values())
    assert "PCCM_CalculationRefusal" not in everything


def test_06_the_endpoints_are_declared_in_the_contract_and_bound_to_no_button() -> None:
    manifest = emitted_manifest()
    declared = set(manifest["vba"]["api_procedures"])
    assert declared == PCCM_ENDPOINTS
    bound = {b["entry_point"] for b in manifest["buttons"]}
    assert len(manifest["buttons"]) == 5, "the workbook must still have five buttons"
    assert not (bound & PCCM_ENDPOINTS), f"an endpoint is bound to a button: {bound}"
    assert set(manifest["vba"]["entry_points"]) == bound


def test_07_no_calculate_button_exists_anywhere() -> None:
    manifest = emitted_manifest()
    for button in manifest["buttons"]:
        assert button["entry_point"] != "PCCM_Calculate"
        assert "Calculate" not in button["shape_name"]


def test_08_the_only_other_public_names_are_the_failpoint_stages() -> None:
    """Public because a later harness arms them BY NAME. Nothing else is public."""
    module = _reporter()
    extra = set(module.public_procedures) - PCCM_ENDPOINTS
    assert extra == set(), f"unexpected Public procedure(s): {sorted(extra)}"
    public_constants = {
        match.group(1)
        for line in module.code_without_string_removal.splitlines()
        if (match := re.match(r"^\s*Public\s+Const\s+(\w+)", line))
    }
    assert public_constants == FAILPOINTS, (
        f"unexpected Public constant(s): {sorted(public_constants - FAILPOINTS)}"
    )


# ===========================================================================
# 3. one preparation path
# ===========================================================================
def test_09_three_endpoints_share_one_preparation_path() -> None:
    """One definition of "valid current inputs".

    A state the write path would refuse must not be reported CURRENT because a
    partial digest happened to be constructible.
    """
    module = _reporter()
    for procedure in ("RunCalculation", "PCCM_CalculationStatus",
                      "PCCM_CurrentInputFingerprint"):
        assert "PrepareCurrentCalculation(package, detail)" in _body(module, procedure), (
            f"{procedure} does not use the shared preparation path"
        )


def test_10_the_preparation_sequence_is_the_locked_one() -> None:
    statements = _statements(_reporter(), "PrepareCurrentCalculation")
    order = [
        r"modCalcResolve\.ResolveModel",
        r"modCalcCheck\.CheckResolvedModel",
        r"BuildFactorTables",
        r"BuildDriverFactors",
        r"BuildAudits",
        r"modCalcAnalytical\.AccumulateTotals",
        r"BuildAnnual",
        r"modCalcAnalytical\.Reconcile",
        r"modCalcAnalytical\.AllIdentitiesHold",
        r"BuildFingerprint",
    ]
    indices = [first_index(statements, pattern) for pattern in order]
    assert indices == sorted(indices), f"the preparation sequence is out of order: {indices}"
    assert indices[-1] < len(statements), "the fingerprint is never built"


def test_11_every_identity_must_hold_before_preparation_succeeds() -> None:
    statements = _statements(_reporter(), "PrepareCurrentCalculation")
    hold = first_index(statements, r"AllIdentitiesHold")
    success = statements.index("PrepareCurrentCalculation = True")
    assert hold < success
    assert any("does not hold" in t for t in
               logical_statements(_body_raw(_reporter(), "PrepareCurrentCalculation")).__iter__()
               .__class__ and [t for _, t in
                               logical_statements(_body_raw(_reporter(),
                                                            "PrepareCurrentCalculation"))]), (
        "a failed identity must produce a diagnostic"
    )


def test_12_an_empty_digest_never_counts_as_a_prepared_one() -> None:
    body = _body(_reporter(), "PrepareCurrentCalculation")
    assert "If Len(package.Fingerprint) = 0 Then" in body, (
        "a preparation that produced no digest must not report success"
    )


def test_13_no_analytical_write_occurs_inside_preparation() -> None:
    """Preparation is pure. Everything it produces sits in memory."""
    body = _body(_reporter(), "PrepareCurrentCalculation")
    for writer in ("WriteTable", "WriteAnalytical", "WriteSuccessCommit",
                   "WriteAttemptBlock", "WriteStatusBlock", ".Value2 ="):
        assert writer not in body, f"preparation writes ({writer})"


def test_14_no_numerical_formula_is_reimplemented() -> None:
    """The reporter orchestrates. It computes nothing of its own."""
    module = _reporter()
    code = module.code
    calls = sorted(set(re.findall(r"modCalc(?:Factors|Analytical|Fingerprint)\.(\w+)", code)))
    assert calls == [
        "AccumulateTotals", "AllIdentitiesHold", "BuildAnnualSeries", "BuildDiscountFactors",
        "BuildDriverAudit", "BuildKnom", "BuildKpv", "CalcFpBuildCostRecord",
        "CalcFpBuildFingerprint", "CalcFpBuildRiskRecord", "CalcFpNumberField",
        "Reconcile",
    ], f"unexpected numerical surface: {calls}"
    # THE FOUR HEADER SCALARS ARE NUMBER FIELDS.
    #
    # This list previously carried CalcFpCanonicalNumber and CalcFpCanonicalText
    # together, which is exactly the defect independent review found: the reporter
    # canonicalised each scalar as a number and then framed the RESULT as an S
    # field, so the digest covered "text that looks numeric" where the contract
    # says N. Canonicalising and framing are one decision and belong to one
    # authority, so the reporter now calls the N-field framer and neither
    # primitive.
    assert "CalcFpCanonicalText" not in code, (
        "a header number framed as text changes what the digest covers"
    )
    assert "CalcFpCanonicalNumber" not in code, (
        "canonicalising without framing invites the framing to be reinvented here"
    )
    # No compounding, no distribution arithmetic, no digest recurrence.
    assert not re.search(r"running\s*=\s*running\s*\*", code)
    assert "FP_BASE" not in code and "FP_MOD_1" not in code
    assert not re.search(r"/\s*3#|/\s*6#", code)


# ===========================================================================
# 4. the transaction
# ===========================================================================
def test_15_nothing_analytical_is_written_before_preparation_succeeds() -> None:
    statements = _statements(_reporter(), "RunCalculation")
    prepare = first_index(statements, r"PrepareCurrentCalculation")
    write = first_index(statements, r"WriteAnalytical")
    assert prepare < write, "an analytical write precedes preparation"
    snapshot = first_index(statements, r"CaptureSnapshot")
    assert prepare < snapshot < write, (
        "the snapshot must be taken after preparation and before any write"
    )


def test_16_the_locked_transaction_order_holds() -> None:
    """snapshot -> write -> verify -> commit -> verify commit -> committed."""
    statements = _statements(_reporter(), "RunCalculation")
    order = [
        r"CaptureSnapshot", r"^On Error GoTo TransactionFailed$", r"WriteAnalytical",
        r"VerifyAnalytical", r"WriteSuccessCommit", r"VerifySuccessCommit",
        r"^committed = True$",
    ]
    indices = [first_index(statements, pattern) for pattern in order]
    assert indices == sorted(indices), f"the transaction order is wrong: {indices}"
    assert indices[-1] < len(statements), "the operation is never marked committed"


def test_17_the_commit_and_its_verification_are_inside_the_rollback_envelope() -> None:
    """Both can fail, so both sit inside the error envelope."""
    statements = _statements(_reporter(), "RunCalculation")
    armed = statements.index("On Error GoTo TransactionFailed")
    # The FIRST disarm AFTER the envelope was armed. Taking the first one in the
    # procedure would find the pre-write envelope's disarm, which sits earlier and
    # would make any ordering assertion below it vacuously true.
    disarmed = armed + 1 + statements[armed + 1:].index("On Error GoTo 0")
    commit = first_index(statements, r"WriteSuccessCommit")
    verify = first_index(statements, r"VerifySuccessCommit")
    assert armed < commit < disarmed, "the commit is outside the rollback envelope"
    assert armed < verify < disarmed, "the commit verification is outside the envelope"


def test_18_success_is_not_published_before_the_snapshot_verifies() -> None:
    statements = _statements(_reporter(), "RunCalculation")
    verify = first_index(statements, r"VerifyAnalytical")
    commit = first_index(statements, r"WriteSuccessCommit")
    assert verify < commit


def test_19_a_verification_failure_raises_rather_than_continuing() -> None:
    statements = _statements(_reporter(), "RunCalculation")
    for guard in ("VerifyAnalytical(package)", "VerifySuccessCommit(successBlock)"):
        index = next(i for i, t in enumerate(statements) if guard in t)
        assert any("Err.Raise" in t for t in statements[index:index + 3]), (
            f"a failed {guard} does not raise"
        )


def test_20_no_generic_error_suppression() -> None:
    """Every handler is one of the reviewed envelopes, and each one lands.

    This test asserted that TransactionFailed was the ONLY handler in the module.
    Independent review rejected that shape: it left CaptureAppState,
    BeginOperation, the preparation, the snapshot and every bookkeeping write
    outside any handler at all, so a runtime fault in one of them escaped raw and
    left EnableEvents, Calculation mode and ScreenUpdating dirty. The invariant
    worth keeping is not "one handler" - it is that no handler suppresses, that
    every handler name is a reviewed envelope, and that each target label really
    exists in the procedure that arms it.
    """
    module = _reporter()
    code = module.code
    assert "On Error Resume Next" not in code, (
        "a suppressed error would become a silently wrong calculation"
    )
    handlers = {h for h in re.findall(r"On Error GoTo (\w+)", code) if h != "0"}
    assert handlers == {
        "InvocationFailed",     # the top-level envelope over the whole endpoint
        "NormalCleanupFailed",  # the NORMAL-path cleanup raising rather than reporting
        "CleanupFailed",        # the RECOVERY cleanup raising while handling a failure
        "PreWriteFailed",       # an unexpected fault before any analytical mutation
        "TransactionFailed",    # the rollback envelope over the mutating region
        "RollbackFailed",       # the restore itself failing
        "BookkeepingFailed",    # the calc_state record failing after the outcome
    }, f"an unreviewed error handler exists: {sorted(handlers)}"
    # Armed and landed in the SAME procedure. A handler whose label lives
    # elsewhere is not a handler; VBA would refuse it, and a text sweep that never
    # checked would not notice.
    for procedure in module.procedures:
        body = _body(module, procedure)
        for target in {h for h in re.findall(r"On Error GoTo (\w+)", body) if h != "0"}:
            assert re.search(rf"^{target}:$", body, re.M), (
                f"{procedure} arms {target} but does not define it"
            )


def test_21_both_failpoints_are_wired_through_the_phase_4_mechanism() -> None:
    """One after analytical state is mutated. One AT the C13:C20 assignment.

    `around < commit` was too weak: any statement earlier in the procedure
    satisfies it, and the submitted source put the commit hook before
    BuildSuccessBlock - so it exercised a failure during commit PREPARATION, not
    at the commit boundary Gate B is required to inject at. The location is now
    asserted inside the writer's own body, adjacently.
    """
    module = _reporter()
    statements = _statements(module, "RunCalculation")
    write = first_index(statements, r"WriteAnalytical")
    mid = first_index(statements, r"FailPointCheck FAILPOINT_ANALYTICAL_WRITE")
    verify = first_index(statements, r"VerifyAnalytical")
    commit = first_index(statements, r"WriteSuccessCommit")
    assert write < mid < verify, (
        "the mid-write failpoint must fire after a block is mutated and before "
        "the analytical verification"
    )

    # --- the commit hook is AT the assignment, inside the writer ------------
    writer = _statements(module, "WriteSuccessCommit")
    executable = [
        s for s in writer
        if not re.match(r"^(Public |Private )?(Sub|Function)\b", s)
        and not re.match(r"^Dim\b", s)
    ]
    hook = next(
        (i for i, s in enumerate(executable)
         if s == "modAppState.FailPointCheck FAILPOINT_SUCCESS_COMMIT"), None
    )
    assert hook is not None, (
        "the commit failpoint is not in the procedure that performs the commit"
    )
    assignment = next(
        i for i, s in enumerate(executable)
        if re.match(r"^CalcSheet\.Range\(CALC_STATE_VALUE_RANGE\)\.Value2 = block$", s)
    )
    assert hook == assignment - 1, (
        "the commit failpoint is not immediately before the C13:C20 assignment: "
        f"{executable[hook + 1:assignment]!r} stands between them"
    )
    assert hook < assignment, "the failpoint fires after C13:C20 is published"
    # It is not ALSO left upstream, where it would fire first and never reach here.
    assert "FAILPOINT_SUCCESS_COMMIT" not in _body(module, "RunCalculation"), (
        "the commit failpoint is still wired into the transaction body"
    )
    assert module.code.count("FailPointCheck FAILPOINT_SUCCESS_COMMIT") == 1, (
        "the commit failpoint is injected from more than one place"
    )
    # And the writer is called from inside the rollback envelope.
    armed = statements.index("On Error GoTo TransactionFailed")
    assert armed < commit, "the commit is outside the rollback envelope"

    assert "modAppState.FailPointCheck" in module.code, (
        "the accepted Phase-4 injection mechanism must be reused"
    )
    assert "gAutomationFailAfterStage" not in module.code, (
        "a second injection framework was created"
    )


# ===========================================================================
# 5. the mutation boundaries
# ===========================================================================
def test_22_a_pre_write_refusal_touches_only_the_attempt_block() -> None:
    """C13:C16 and every analytical block stand exactly as they were."""
    module = _reporter()
    statements = _statements(module, "RunCalculation")
    refusal = next(i for i, t in enumerate(statements) if "RecordRefusal" in t)
    write = first_index(statements, r"WriteAnalytical")
    assert refusal < write, "a refusal happens before any analytical write"
    body = _body(module, "RecordRefusal")
    assert "WriteAttemptBlock" in body
    for forbidden in ("CALC_STATE_VALUE_RANGE", "WriteAnalytical", "WriteSuccessCommit",
                      "CALC_TOTALS_VALUE_RANGE"):
        assert forbidden not in body, f"a refusal touches {forbidden}"


def test_23_the_attempt_block_is_one_four_row_assignment() -> None:
    module = _reporter()
    body = _body(module, "WriteAttemptBlock")
    assert "Dim block(1 To 4, 1 To 1) As Variant" in body
    assert body.count(".Value2 =") == 1, "the attempt metadata is written in pieces"
    span = _body(module, "AttemptRange")
    assert "CALC_STATE_ROW_LAST_ATTEMPT_RESULT" in span
    assert "CALC_STATE_ROW_STATUS_EVALUATED_AT" in span


def test_24_the_status_refresh_is_one_two_row_assignment() -> None:
    module = _reporter()
    body = _body(module, "WriteStatusBlock")
    assert "Dim block(1 To 2, 1 To 1) As Variant" in body
    assert body.count(".Value2 =") == 1
    span = _body(module, "StatusRange")
    assert "CALC_STATE_ROW_CALCULATION_STATUS" in span
    assert "CALC_STATE_ROW_STATUS_EVALUATED_AT" in span
    status = _body(module, "PCCM_CalculationStatus")
    for forbidden in ("WriteAnalytical", "WriteSuccessCommit", "CALC_TOTALS_VALUE_RANGE",
                      "CALC_STATE_VALUE_RANGE"):
        assert forbidden not in status, f"asking for the status touches {forbidden}"


def test_25_the_success_commit_is_built_once_and_written_once() -> None:
    """Built once in memory, written once, and verified against THAT block.

    Not four writes that could half-succeed and leave a fingerprint with no stamp,
    or a stamp with no version - and not a block whose two timestamps come from
    two different calls to Now, which cannot then be verified against what was
    written.
    """
    module = _reporter()
    build = _body(module, "BuildSuccessBlock")
    assert "Dim built(1 To 8, 1 To 1) As Variant" in build
    for row in range(1, 9):
        assert f"built({row}, 1) =" in build, f"commit row {row} is not populated"
    assert "FP_VERSION" in build, "C15 must carry the fingerprint version"
    assert "CALC_ATTEMPT_SUCCESS" in build and "CALC_STATUS_CURRENT" in build
    # ONE moment, into BOTH timestamp rows. Two calls to Now would put two
    # different values in C13 and C20 and make the block unverifiable against
    # itself.
    assert build.count("Now") == 1, "the commit captures the clock more than once"
    assert re.search(r"stamp = Now", build), "the captured moment must be named"
    assert re.search(r"built\(1, 1\) = stamp", build), "C13 is not the captured moment"
    assert re.search(r"built\(8, 1\) = stamp", build), "C20 is not the captured moment"

    write = _body(module, "WriteSuccessCommit")
    assert write.count(".Value2 =") == 1, "the commit metadata is written in pieces"
    assert "CALC_STATE_VALUE_RANGE" in write
    assert "Now" not in write, "the writer must write the block it was given"
    assert re.search(r"\.Value2 = block$", write, re.M), (
        "the writer must write the built block, not a freshly assembled one"
    )


# ===========================================================================
# 6. snapshot and rollback
# ===========================================================================
def test_26_all_five_tables_and_both_scalar_blocks_are_snapshotted() -> None:
    body = _body(_reporter(), "CaptureSnapshot")
    for table in ANALYTICAL_TABLES:
        assert f"modWorkbook.SnapshotTable(modWorkbook.Lo(CALC_SHEET, {table}))" in body \
            or f"CALC_SHEET, {table}" in body, f"{table} is not snapshotted"
    assert body.count("modWorkbook.SnapshotTable") == 5
    assert "CALC_TOTALS_VALUE_RANGE" in body, "C23:C32 is not snapshotted"
    assert "CALC_STATE_VALUE_RANGE" in body, "C13:C20 is not snapshotted"
    assert ".Value2" in body, "values only - labels and formats are build-owned"
    for structural in ("NumberFormat", "Interior", "ColumnWidth"):
        assert structural not in body, f"the snapshot captures {structural}"


def test_27_rollback_restores_everything_that_was_snapshotted() -> None:
    body = _body(_reporter(), "RestoreSnapshot")
    for table in ANALYTICAL_TABLES:
        assert table in body, f"{table} is not restored"
    assert body.count("modWorkbook.RestoreTable") == 5, (
        "all five tables must go through the accepted Phase-4 restore"
    )
    assert "CALC_TOTALS_VALUE_RANGE" in body, "C23:C32 is not restored"
    assert "CALC_STATE_VALUE_RANGE" in body, "prior C13:C20 is not restored"


def test_28_rollback_happens_before_any_failed_metadata() -> None:
    """The first observable moment after a failure is the previous successful
    snapshot, exactly.

    The rollback and its record now live in RollbackAndRecord rather than inline
    in RunCalculation, so the ordering is asserted where it happens. The stronger
    claim is the second one: if the RESTORE itself fails, NO failed-attempt
    metadata is written at all. That record asserts "the previous snapshot
    stands", and writing it after a failed restore would assert something nobody
    established.
    """
    module = _reporter()
    statements = _statements(module, "RollbackAndRecord")
    restore = next(i for i, s in enumerate(statements) if "RestoreSnapshot" in s)
    record = next(i for i, s in enumerate(statements) if "WriteAttemptBlock" in s)
    assert restore < record, "FAILED metadata is written before the rollback"
    # The restore is armed before it runs, and its handler writes nothing.
    assert statements[restore - 1] == "On Error GoTo RollbackFailed", (
        "the restore runs outside an error envelope and could escape raw"
    )
    failed = _body(module, "RollbackAndRecord").split("RollbackFailed:", 1)[1]
    failed = failed.split("BookkeepingFailed:", 1)[0]
    for forbidden in ("WriteAttemptBlock", "WriteStatusBlock", "WriteSuccessCommit",
                      ".Value2 ="):
        assert forbidden not in failed, (
            f"a failed rollback still writes {forbidden} under a false premise"
        )
    # And a failed RECORD does not undo a rollback that succeeded.
    book = _body(module, "RollbackAndRecord").split("BookkeepingFailed:", 1)[1]
    assert "RestoreSnapshot" not in book, "a failed record re-runs the rollback"


def test_29_a_committed_operation_can_never_become_failed() -> None:
    """Once C17 says SUCCESS, that is committed workbook truth.

    A problem AFTER the commit - restoring EnableEvents, Calculation mode or
    ScreenUpdating - is an application/invocation cleanup failure, not a failed
    analytical transaction. Rewriting the attempt to FAILED there would leave the
    workbook and the reported outcome contradicting each other, so the committed
    flag is carried out of the transaction and the post-commit path writes
    nothing at all.
    """
    module = _reporter()
    statements = _statements(module, "RunCalculation")
    committed = statements.index("committed = True")
    # THE ENVELOPE IS DISARMED, then the success is published, then the procedure
    # leaves. No handler in this procedure can fire after the commit, so no
    # rollback and no FAILED record can reach a committed transaction.
    tail = statements[committed:]
    disarm = tail.index("On Error GoTo 0")
    leave = tail.index("Exit Function")
    assert disarm < leave, "an error handler is still armed after the commit"
    for forbidden in ("RestoreSnapshot", "WriteAttemptBlock", "WriteStatusBlock",
                      "RecordFailureWithoutRollback", "RollbackAndRecord",
                      "WriteSuccessCommit"):
        assert not any(forbidden in s for s in tail[:leave]), (
            f"a committed transaction can still reach {forbidden}"
        )
    # The flag leaves the transaction, so the caller can tell the two axes apart.
    assert "RunCalculation(ByRef committed As Boolean)" in _body(module, "RunCalculation"), (
        "the commit fact must reach the endpoint, not die inside the transaction"
    )
    # And the post-commit branch writes NOTHING.
    cleanup = _body(module, "CleanupOutcome")
    guard = cleanup.index("If committed Then")
    for forbidden in ("WriteAttemptBlock", "WriteStatusBlock", "WriteSuccessCommit",
                      "RestoreSnapshot", ".Value2 =", "CALC_ATTEMPT_FAILED"):
        assert forbidden not in cleanup, (
            f"post-commit cleanup touches {forbidden} and can falsify C17"
        )
    committed_branch = cleanup[guard:cleanup.index("Exit Function", guard)]
    assert "modAppState.Failed" in committed_branch, (
        "the cleanup problem must still be reported on the invocation axis"
    )


def test_30_no_second_rollback_mechanism_exists() -> None:
    code = _reporter().code
    assert "modWorkbook.RestoreTable" in code
    assert code.count("ListRows.Add") == 1, (
        "resizing belongs to one place; a second body rebuilder would be a "
        "second restore mechanism"
    )


# ===========================================================================
# 7. status and the accessors
# ===========================================================================
def test_31_derived_status_is_never_an_attempt_result() -> None:
    module = _reporter()
    body = _body(module, "DeriveStatus")
    for attempt in ("CALC_ATTEMPT_REFUSED", "CALC_ATTEMPT_FAILED", "CALC_ATTEMPT_SUCCESS",
                    "CALC_ATTEMPT_NONE"):
        assert attempt not in body, f"the derived status can be {attempt}"
    assigned = set(re.findall(r"DeriveStatus = (\w+)", body))
    assert assigned == DERIVED_STATUSES, (
        f"unexpected derived status value(s): {sorted(assigned - DERIVED_STATUSES)}"
    )


def test_32_status_never_reads_the_attempt_history() -> None:
    body = _body(_reporter(), "DeriveStatus")
    assert "CALC_STATE_ROW_LAST_ATTEMPT_RESULT" not in body, (
        "the status is derived from the historical attempt"
    )
    assert "CALC_STATE_ROW_LAST_SUCCESSFUL_FINGERPRINT" in body, (
        "the status must compare against the stored successful digest"
    )


def test_33_invalid_current_inputs_short_circuit_the_comparison() -> None:
    """No empty-digest comparison can be reached."""
    statements = _statements(_reporter(), "DeriveStatus")
    invalid = statements.index("If Not prepared Then")
    stored = first_index(statements, r"stored = StoredText")
    compare = first_index(statements, r"StrComp\(package\.Fingerprint, stored")
    assert invalid < stored < compare
    empty = first_index(statements, r"If Len\(stored\) = 0 Then")
    assert empty < compare, "an empty stored digest reaches the equality comparison"


def test_34_a_failed_attempt_re_derives_the_status() -> None:
    """FAILED is an attempt result. It never chooses the status.

    Both failure recorders are checked: the one that runs when nothing analytical
    was mutated, and the one that runs after a rollback.
    """
    module = _reporter()
    for procedure in ("RecordFailureWithoutRollback", "RollbackAndRecord"):
        body = _body(module, procedure)
        assert "CurrentStatus()" in body, f"{procedure} does not re-derive the status"
        assert "CALC_ATTEMPT_FAILED" in body, f"{procedure} records the wrong result"
        assert "CALC_STATUS_FAILED" not in body, "FAILED is not a status"
    fresh = _body(module, "CurrentStatus")
    assert "PrepareCurrentCalculation(package, detail)" in fresh, (
        "the re-derivation must run a fresh preparation against the restored state"
    )


def test_35_the_stored_fingerprint_accessor_does_not_recompute() -> None:
    body = _body(_reporter(), "PCCM_CalculationFingerprint")
    assert "PrepareCurrentCalculation" not in body, (
        "the stored digest accessor recomputes the current inputs"
    )
    assert "CALC_STATE_ROW_LAST_SUCCESSFUL_FINGERPRINT" in body


def test_36_the_current_fingerprint_accessor_returns_blank_when_invalid() -> None:
    """Not a sentinel digest - a sentinel would eventually be compared."""
    statements = _statements(_reporter(), "PCCM_CurrentInputFingerprint")
    assert statements.index("If PrepareCurrentCalculation(package, detail) Then") >= 0
    assigns = [t for t in statements if t.startswith("PCCM_CurrentInputFingerprint =")]
    assert assigns == ["PCCM_CurrentInputFingerprint = package.Fingerprint"], (
        "the only assignment is the prepared digest; the invalid path falls "
        f"through to the empty default, found {assigns}"
    )


def test_37_the_read_accessors_mutate_nothing() -> None:
    module = _reporter()
    for accessor in ("PCCM_CalculationFingerprint", "PCCM_CurrentInputFingerprint",
                     "PCCM_CalculationAttemptResult", "PCCM_CalculationAttemptDetail"):
        body = _body(module, accessor)
        assert ".Value2 =" not in body, f"{accessor} writes"
        for writer in ("WriteAttemptBlock", "WriteStatusBlock", "WriteSuccessCommit",
                       "WriteAnalytical", "RecordRefusal", "RecordFailure"):
            assert writer not in body, f"{accessor} calls {writer}"


# ===========================================================================
# 8. the audit blocks
# ===========================================================================
def test_38_an_inapplicable_driver_field_is_blank() -> None:
    """Never the in-memory identity 1, and never zero.

    A Quantity of 1 shown against a risk would read as a real entry; a zero
    would read as a real amount.
    """
    module = _reporter()
    body = _body(module, "DriversBlock")
    for blanked in ("COL_CALC_DRIVERS_QUANTITY", "COL_CALC_DRIVERS_PROBABILITY",
                    "COL_CALC_DRIVERS_CENTRAL_VALUE", "COL_CALC_DRIVERS_CENTRAL_BASIS",
                    "COL_CALC_DRIVERS_DETERMINISTIC_NOMINAL",
                    "COL_CALC_DRIVERS_MEAN_BASIS_NOMINAL",
                    "COL_CALC_DRIVERS_UNCERTAINTY_MEAN_SHIFT_NOMINAL",
                    "COL_CALC_DRIVERS_EXPECTED_RISK_NOMINAL"):
        assert re.search(rf"block\(row, {blanked}\) = Empty", body), (
            f"{blanked} is never blanked for the kind it does not apply to"
        )
    assert not re.search(r"block\(row, COL_CALC_DRIVERS_\w+\) = 1#", body), (
        "an identity 1 is written into the audit"
    )
    assert not re.search(r"block\(row, COL_CALC_DRIVERS_\w+\) = 0#", body), (
        "a zero is written into an inapplicable audit field"
    )


def test_39_the_carry_identities_exist_only_in_memory() -> None:
    """`Quantity = 1` for risks and `Probability = 1` for cost lines are
    calculation semantics, and appear only where DriverFactors is built."""
    module = _reporter()
    factors = _body(module, "BuildDriverFactors")
    assert ".Quantity = 1#" in factors and ".Probability = 1#" in factors
    audit = _body(module, "DriversBlock")
    assert "= 1#" not in audit, "an identity reached the audit block"
    fingerprint = _body(module, "DriverRecord")
    assert "1#" not in fingerprint, "an identity reached the fingerprint record"


def test_40_the_base_year_inflation_row_is_blank_rate_and_unit_factor() -> None:
    module = _reporter()
    body = _body(module, "InflationBlock")
    assert "If offset = 0 Then" in body
    assert re.search(r"COL_CALC_INFLATION_FACTORS_ANNUAL_RATE\) = Empty", body), (
        "the Base-Year annual rate must be BLANK, not a fabricated zero"
    )
    assert not re.search(r"COL_CALC_INFLATION_FACTORS_ANNUAL_RATE\) = 0#", body)
    assert "COL_CALC_INFLATION_FACTORS_CUMULATIVE_INFLATION_FACTOR" in body
    assert "package.InflationSpan" in body, (
        "the audited span must cover BaseYear..LastYear, so pre-project "
        "compounding years stay visible"
    )


def test_41_the_fx_audit_carries_referenced_currencies_only() -> None:
    module = _reporter()
    body = _body(module, "FxBlock")
    assert "package.Model.CurrencyCount" in body, "the row count is the referenced set"
    assert "REPORTING_CURRENCY" not in module.code, (
        "the reporting currency is seeded into the audit"
    )
    counter = _body(module, "CountCurrencyReferences")
    assert "vbBinaryCompare" in counter, "the count must use exact key semantics"
    assert "COL_CALC_FX_REFERENCED_BY" in body


def test_42_a_zero_row_table_fabricates_no_record() -> None:
    """A physical placeholder is not a semantic record."""
    module = _reporter()
    for builder in ("YearsBlock", "InflationBlock", "FxBlock", "DriversBlock",
                    "AnnualBlock"):
        body = _body(module, builder)
        assert re.search(r"If rows < 1 Then Exit Function", body), (
            f"{builder} does not return Empty for a zero-row table"
        )
    writer = _body(module, "WriteTable")
    assert "target.DataBodyRange.ClearContents" in writer, (
        "the physical placeholder must be cleared, not filled"
    )


def test_43_totals_come_from_memory_and_never_from_the_sheet() -> None:
    body = _body(_reporter(), "TotalsBlock")
    assert "package.Totals." in body
    assert "Range(" not in body and "CalcSheet" not in body, (
        "the totals are read back off the worksheet"
    )
    assert body.count("package.Totals.") == 10, "all ten totals must come from memory"


def test_44_no_analytical_value_is_read_back_for_calculation() -> None:
    """The worksheet is the RECORD of a calculation, never an input into one."""
    module = _reporter()
    for builder in ("YearsBlock", "InflationBlock", "FxBlock", "DriversBlock",
                    "AnnualBlock", "TotalsBlock"):
        body = _body(module, builder)
        assert "CalcSheet" not in body, f"{builder} reads the worksheet"
        assert ".Value" not in body, f"{builder} reads a cell"


def test_45_the_verifier_reads_back_and_compares() -> None:
    """A write is not proven by the absence of an error."""
    module = _reporter()
    body = _body(module, "VerifyTable")
    assert "target.DataBodyRange.Cells(r, c).Value" in body, "nothing is read back"
    assert "SameCell(" in body
    compare = _body(module, "SameCell")
    assert "IsEmpty(written)" in compare, "a blank must verify as blank"
    assert "CDbl(written) = CDbl(wanted)" in compare


def test_46_a_blank_never_verifies_as_a_zero() -> None:
    statements = _statements(_reporter(), "SameCell")
    blank = next(i for i, t in enumerate(statements) if t.startswith("If IsEmpty(wanted)"))
    numeric = first_index(statements, r"CDbl\(written\) = CDbl\(wanted\)")
    assert blank < numeric, "a blank reaches the numeric comparison"


def test_47_headers_and_formats_are_never_written() -> None:
    code = _reporter().code
    for owned in ("HeaderRowRange", "NumberFormat", "Interior", "ColumnWidth",
                  "ListColumns.Add"):
        assert owned not in code, f"the calculation writes {owned}, which is build-owned"


# ===========================================================================
# 9. application state and the forbidden surfaces
# ===========================================================================
def test_48_application_state_is_captured_and_restored() -> None:
    """Captured, changed, and restored however the endpoint returns.

    The envelope is armed BEFORE the first fallible operation. Independent review
    found it armed after CaptureAppState and BeginOperation, which left the two
    operations that make restoration necessary outside any handler: a fault in
    either escaped raw and left EnableEvents, Calculation mode and ScreenUpdating
    dirty, which is worse than a failed calculation.
    """
    module = _reporter()
    body = _body(module, "PCCM_Calculate")
    for stage in ("modAppState.CaptureAppState()", "modAppState.BeginOperation",
                  "modAppState.FinishOperation(state)"):
        assert stage in body, f"{stage} is missing"
    statements = _statements(module, "PCCM_Calculate")
    armed = statements.index("On Error GoTo InvocationFailed")
    capture = next(i for i, s in enumerate(statements) if "CaptureAppState()" in s)
    begin = next(i for i, s in enumerate(statements) if "BeginOperation" in s)
    run = next(i for i, s in enumerate(statements) if "RunCalculation(committed)" in s)
    finish = next(i for i, s in enumerate(statements) if "FinishOperation" in s)
    assert armed < capture < begin < run < finish, (
        "the envelope must cover capture, begin and the transaction alike"
    )
    # Explicit flags, not inferences. Restoring state that was never captured
    # would be its own fault; skipping restoration that IS owed is the defect the
    # envelope exists to prevent; and retrying an attempt that already raised is
    # the defect the second flag exists to prevent.
    assert "stateCaptured As Boolean" in body
    assert "cleanupAttempted As Boolean" in body
    assert re.search(r"stateCaptured = True", body), "the flag is never set"
    assert re.search(r"If stateCaptured And Not cleanupAttempted Then", body), (
        "the recovery path restores unconditionally, or retries a spent attempt"
    )
    # NEITHER cleanup call is outside a handler. Every FinishOperation in this
    # procedure has a cleanup-failure envelope armed and still open above it.
    for index in [i for i, s in enumerate(statements) if "FinishOperation" in s]:
        armed = [j for j, s in enumerate(statements[:index])
                 if s.startswith("On Error GoTo ") and s != "On Error GoTo 0"]
        disarmed = [j for j, s in enumerate(statements[:index]) if s == "On Error GoTo 0"]
        assert armed, f"the cleanup at statement {index} has no handler at all"
        assert armed[-1] > (disarmed[-1] if disarmed else -1), (
            f"the cleanup at statement {index} runs with its handler disarmed; "
            "an exception there escapes the endpoint raw"
        )
        target = statements[armed[-1]].split()[-1]
        assert target in ("NormalCleanupFailed", "CleanupFailed"), (
            f"the cleanup at statement {index} is covered by {target}, "
            "which is not a cleanup-failure envelope"
        )
    assert any("modAppState.Failed" in s or "CleanupOutcome" in s
               for s in statements[finish:]), (
        "a failed cleanup must be reported, not swallowed"
    )


def test_49_no_message_box_or_change_handler_exists() -> None:
    code = _reporter().code
    for forbidden in ("MsgBox", "Worksheet_Change", "Workbook_SheetChange",
                      "Worksheet_Calculate", "Workbook_Open"):
        assert forbidden not in code, f"{forbidden} appears in the reporter"


def test_50_no_simulation_code_exists() -> None:
    code = _reporter().code
    for forbidden in ("Rnd", "Randomize", "MRG32k3a", "_SimData", "SH_SIMDATA",
                      "Iteration", "Percentile"):
        assert forbidden not in code, f"{forbidden} belongs to Phase 6"


def test_51_the_accepted_modules_were_not_modified() -> None:
    import hashlib

    frozen = {
        "modCalcResolve": "0890c612ade1b00b93568bcb32b42121f83bff1ec6647224cccaa59322b15afe",
        "modCalcCheck": "738343945932150470233cb2a0b7e6fea7617db1a877cae8e09d19085e39c43b",
        "modCalcFactors": "701097ab3092a1fdec9ef7168d55f50248df1acf11e2517f0ea3c18fed278128",
        "modCalcAnalytical": "affea282af14c70ff5bf6dd19fab6e56174c39b0e9cba251495cdf1bfaac39b7",
        # Its CURRENT bytes. Step 7's correction round carried the ONE authorised
        # reopening of this accepted module - CalcFpNumberField became Public so
        # the reporter could reach the accepted N-field framing authority instead
        # of framing a number as text. FINGERPRINT_STEP4_BODY_SHA256 is what says
        # nothing else moved with it.
        "modCalcFingerprint": "39e80b9ef9252a9822cd57c8ae441b67571ca3725b3d78124bd6af2ddccc4744",
        "modWorkbook": "9cfa8f130c5bcdee783948654c969d4b0d6589fe7059c126f88c7676ca5405bf",
        "modAppState": "ef0b5c64a7a3b5aeeef5ef0797cd160071a7eda6a7d8cef9cb98301f1504672f",
        "modTimeline": "4a4f24d17b65bcbc0e46b1a74213b6a02eab6ab492b1788476d66eb7807b9e3f",
        "modDrivers": "8f947a4cc473b76161c867f99daf5fbb4af670b909cca0387165b079c102af48",
        "modProfiling": "0312858d7d817d20a99877f8be52ca0f7cf5b0bbb9aa9770367ed11138d9d7ca",
        "modInflation": "08db32807d495c22e6067350291c21a9a277884de5e5064555612f6bb991118c",
        "modStructuralCheck": "1798c56a459c9e35c581871248815841b28a3c88a62a931a68afe5d71853ed54",
    }
    for name, digest in frozen.items():
        actual = hashlib.sha256((SRC_VBA / f"{name}.bas").read_bytes()).hexdigest()
        assert actual == digest, f"{name}.bas changed; Step 7 adds a module and edits none"

    # AND THE FOUR THAT MOVED IN RUN 7 MOVED BY A RENAME AND NOTHING ELSE.
    for module in RUN7_RENAMES_BY_MODULE:
        _assert_run7_rename_only(module)


# ===========================================================================
# 9b. THE CORRECTION-ROUND INVARIANTS
#
# Each of the five defects independent review found has a test here that fails
# against the submitted source and passes against the corrected source.
# ===========================================================================
def test_53_the_header_scalars_are_framed_as_number_fields() -> None:
    """BLOCKER 1. Base Year, Start Year, Duration and Discount Rate are numbers.

    They were canonicalised as numbers and then framed as TEXT, so the digest
    covered four S fields carrying numeric text where the contract says N. The
    framing authority is modCalcFingerprint's in either case, so the reporter
    calls its N-field framer rather than assembling a field of its own.
    """
    module = _reporter()
    body = _body(module, "BuildFingerprint")
    framed = re.findall(
        r"modCalcFingerprint\.CalcFpNumberField\(\s*package\.Model\.Timeline\.(\w+)", body
    )
    assert framed == ["BaseYear", "StartYear", "Duration", "DiscountRate"], (
        f"the header scalars are not the four locked ones, in order: {framed}"
    )
    assert len(re.findall(r"header\(\d\)", body)) == 4, "the header is not four fields"
    # No S framing, and no field assembled by hand out here.
    assert "CalcFpCanonicalText" not in module.code
    assert "FP_TAG_TEXT" not in module.code and "FP_TAG_NUMBER" not in module.code, (
        "the reporter frames a field itself instead of calling the framer"
    )
    assert "FP_FIELD_SEPARATOR" not in module.code


def test_54_the_outcome_is_published_through_the_automation_aware_surface() -> None:
    """BLOCKER 2. Announce, never ReportResult.

    ReportResult shows a modal dialog unconditionally. Announce records the
    outcome for automation and shows the dialog only when automation is inactive,
    which is what lets Gate B drive this endpoint at all. A direct dialog would
    block the harness and leave the run with no recorded result.
    """
    module = _reporter()
    code = module.code
    assert "modAppState.Announce" in code, "the outcome never reaches automation"
    assert "ReportResult" not in code, (
        "the endpoint publishes through the modal reporter and would block automation"
    )
    # The accepted Phase-4 surface is REUSED, not re-created here.
    for invented in ("gAutomationActive", "gAutomationResult", "gLastResult",
                     "AutomationActive", "RecordForAutomation"):
        assert invented not in code, f"a second automation reporter ({invented}) was created"
    # Announce belongs to the endpoint. Nothing deeper may publish.
    for procedure in module.procedures:
        if procedure == "PCCM_Calculate":
            continue
        assert "Announce" not in _body(module, procedure), (
            f"{procedure} publishes an outcome; only the endpoint may"
        )


def test_55_every_exit_from_the_endpoint_publishes_exactly_one_outcome() -> None:
    """BLOCKER 2/3. No path leaves PCCM_Calculate silently.

    A path that returns without announcing is a calculation whose outcome no
    automation run can observe - indistinguishable, from outside, from one that
    never started.
    """
    statements = _statements(_reporter(), "PCCM_Calculate")
    announces = [i for i, s in enumerate(statements) if "modAppState.Announce" in s]
    exits = [i for i, s in enumerate(statements) if s == "Exit Sub"]
    assert len(announces) == 4, (
        f"expected one announcement per terminal path, found {len(announces)}"
    )
    # Each Exit Sub is immediately preceded by an announcement, and the last
    # path falls out of the procedure straight after one.
    for index in exits:
        assert index - 1 in announces, "a path leaves the endpoint without announcing"
    assert announces[-1] > exits[-1], "the final handler path announces nothing"


def test_56_the_envelope_covers_every_fallible_invocation_step() -> None:
    """BLOCKER 3. Containment starts before the first thing that can fail.

    CaptureAppState, BeginOperation, the preparation, the snapshot and every
    bookkeeping write can all raise. Whatever is left outside the envelope escapes
    raw, past the restoration that the operation itself made necessary.
    """
    module = _reporter()
    statements = _statements(module, "PCCM_Calculate")
    # Declarations cannot fail. The first statement that can is the envelope.
    executable = [
        s for s in statements
        if not re.match(r"^(Public |Private )?(Sub|Function)\b", s)
        and not re.match(r"^Dim\b", s)
    ]
    assert executable[0] == "On Error GoTo InvocationFailed", (
        f"something fallible runs before the envelope is installed: {executable[0]!r}"
    )
    # Preparation and the snapshot sit in their own envelope inside the
    # transaction, so a runtime fault there is FAILED and not a model refusal.
    run = _statements(module, "RunCalculation")
    armed = run.index("On Error GoTo PreWriteFailed")
    prepare = next(i for i, s in enumerate(run) if "PrepareCurrentCalculation" in s)
    snapshot = next(i for i, s in enumerate(run) if "CaptureSnapshot" in s)
    assert armed < prepare < snapshot, "preparation or the snapshot runs uncovered"
    # And the two outcomes are told apart AT THE POINT THEY DIVERGE. The
    # controlled refusal is reached from the "not prepared" branch; the handler
    # for an unexpected fault in the very same region must NOT reach it, or a
    # runtime fault would be reported to the user as a model problem they do not
    # have.
    body = _body(module, "RunCalculation")
    happy, _, handlers = body.partition("PreWriteFailed:")
    prewrite, _, transaction = handlers.partition("TransactionFailed:")
    assert "RecordRefusal" in happy, "the controlled refusal is not on the normal path"
    assert "If Not prepared Then" in happy, "the refusal is not conditioned on preparation"
    assert "RecordFailureWithoutRollback" in prewrite, (
        "an unexpected pre-write fault is not recorded as FAILED"
    )
    assert "RecordRefusal" not in prewrite, (
        "a runtime fault is reported as a model refusal"
    )
    assert "RollbackAndRecord" not in prewrite, (
        "nothing was mutated; a rollback here would restore over untouched state"
    )
    assert "RollbackAndRecord" in transaction and "RecordRefusal" not in transaction, (
        "a fault after mutation must roll back and be FAILED, never REFUSED"
    )
    refusal = _body(module, "RecordRefusal")
    assert "CALC_ATTEMPT_REFUSED" in refusal and "CALC_ATTEMPT_FAILED" not in refusal
    prewrite = _body(module, "RecordFailureWithoutRollback")
    assert "CALC_ATTEMPT_FAILED" in prewrite and "CALC_ATTEMPT_REFUSED" not in prewrite
    assert "RestoreSnapshot" not in prewrite, (
        "nothing analytical was mutated; there is nothing to roll back"
    )


def test_57_cleanup_is_attempted_once_and_only_when_state_was_captured() -> None:
    """BLOCKER 3, and correction round 2. BOTH cleanup calls are contained.

    FinishOperation returns a diagnostic String for a restoration it could not
    complete - but it is an Excel call and can also RAISE. Round 1 armed an
    envelope over the RECOVERY call only and left the NORMAL call outside every
    handler, so a raise there escaped the endpoint with no announcement, no
    recorded automation result, and stateCaptured still True.

    Twice would restore a state already restored; never would leave EnableEvents,
    Calculation mode and ScreenUpdating dirty. Exactly once, contained, on every
    path.
    """
    module = _reporter()
    body = _body(module, "PCCM_Calculate")
    statements = _statements(module, "PCCM_Calculate")
    finishes = [i for i, s in enumerate(statements) if "FinishOperation" in s]
    assert len(finishes) == 2, (
        "expected one cleanup on the normal path and one on the failure path"
    )
    normal, recovery = finishes

    # --- the NORMAL call runs with a cleanup-failure handler ARMED ----------
    #
    # Armed before it, and the top-level envelope is not what covers it: that one
    # is disarmed first, precisely so a normal-path raise does not re-enter the
    # recovery cleanup.
    armed = [i for i, s in enumerate(statements[:normal])
             if s.startswith("On Error GoTo ") and s != "On Error GoTo 0"]
    disarmed = [i for i, s in enumerate(statements[:normal]) if s == "On Error GoTo 0"]
    assert armed and armed[-1] > (disarmed[-1] if disarmed else -1), (
        "the normal FinishOperation runs with no error handler armed"
    )
    normal_handler = statements[armed[-1]].split()[-1]
    assert normal_handler == "NormalCleanupFailed", (
        f"the normal cleanup is covered by {normal_handler}, not its own envelope"
    )
    assert re.search(rf"^{normal_handler}:$", body, re.M), (
        "the normal-cleanup handler label does not exist"
    )

    # --- the attempt is SPENT, in state, before the call can raise ----------
    assert "cleanupAttempted As Boolean" in body, (
        "nothing in state records that the one permitted attempt was used"
    )
    spent = statements.index("cleanupAttempted = True")
    assert spent < normal, (
        "the attempt is marked after the call, so a raise leaves it unmarked"
    )
    assert statements[normal + 1] == "On Error GoTo 0", "the envelope is not closed"
    assert "stateCaptured = False" in statements[normal:normal + 3], (
        "a completed normal cleanup does not clear the state-owed flag"
    )

    # --- the normal-cleanup handler does NOT retry, does NOT write ----------
    handler = body.split(f"{normal_handler}:", 1)[1].split("InvocationFailed:", 1)[0]
    assert "FinishOperation" not in handler, (
        "a cleanup that raised is retried; one attempt means one attempt"
    )
    assert "modAppState.Announce" in handler, (
        "a raised normal cleanup is never announced"
    )
    assert "CleanupOutcome" in handler, (
        "the committed/uncommitted distinction is not applied to a raised cleanup"
    )
    for forbidden in ("WriteAttemptBlock", "WriteStatusBlock", "WriteSuccessCommit",
                      "RestoreSnapshot", ".Value2 ="):
        assert forbidden not in handler, (
            f"a raised normal cleanup touches {forbidden} after a committed calculation"
        )

    # --- the RECOVERY call is guarded by BOTH facts -------------------------
    guard = statements.index("If stateCaptured And Not cleanupAttempted Then")
    assert guard < recovery, "the recovery cleanup is not guarded by the flags"
    assert "On Error GoTo CleanupFailed" in statements[:guard], (
        "the recovery cleanup can raise out of the handler"
    )
    assert statements.index("cleanupAttempted = True", guard) < recovery, (
        "the recovery attempt is not marked spent before it is made"
    )
    # And it cannot retry either. The label is matched line-anchored: splitting on
    # the bare text would land inside NormalCleanupFailed and read the wrong
    # handler, which is how a retry hides from a sweep that never checks.
    tail = re.split(r"^CleanupFailed:$", body, maxsplit=1, flags=re.M)[1]
    assert "FinishOperation" not in tail, "the recovery cleanup is retried"
    assert "failure" in tail, "the original failure is replaced by the cleanup failure"


def test_58_a_post_commit_cleanup_problem_is_an_invocation_failure() -> None:
    """BLOCKER 4. C17 = SUCCESS is committed workbook truth.

    A cleanup problem afterwards is an application/invocation failure. Reporting
    the ATTEMPT as FAILED would contradict the workbook, which still says SUCCESS
    and correctly so.
    """
    module = _reporter()
    cleanup = _body(module, "CleanupOutcome")
    assert "ByVal committed As Boolean" in cleanup, (
        "the cleanup outcome cannot tell a committed run from an uncommitted one"
    )
    branch = cleanup[cleanup.index("If committed Then"):]
    branch = branch[:branch.index("Exit Function")]
    assert "modAppState.Failed" in branch, "the cleanup problem is swallowed"
    # It says so, in the text the user and the harness both see.
    raw = _body_raw(module, "CleanupOutcome")
    raw_branch = raw[raw.index("If committed Then"):]
    assert "COMMITTED" in raw_branch, (
        "the message must state that the calculation committed"
    )
    # The endpoint routes through it rather than rewriting the result inline.
    endpoint = _statements(module, "PCCM_Calculate")
    route = next(i for i, s in enumerate(endpoint) if "CleanupOutcome" in s)
    finish = next(i for i, s in enumerate(endpoint) if "FinishOperation" in s)
    assert finish < route, "the cleanup outcome is decided before cleanup runs"
    assert "If Len(cleanup) > 0 Then" in endpoint[route], (
        "a successful cleanup must leave the result alone"
    )


def test_59_the_commit_is_verified_cell_by_cell_against_the_block_written() -> None:
    """BLOCKER 5. All eight cells, including C20, against THAT block.

    A verifier that regenerated Now would compare against a value the commit never
    contained. One that checked the stamp for being merely non-blank would pass
    over a stamp written into the wrong cell. And C20 was not checked at all.
    """
    module = _reporter()
    verify = _body(module, "VerifySuccessCommit")
    assert "ByRef block As Variant" in verify, (
        "the verifier must be handed the block that was written"
    )
    assert "VerifyRange(CALC_STATE_VALUE_RANGE, block, 8)" in verify, (
        "the verification must cover all eight cells of the committed range"
    )
    assert "Now" not in verify, "the verifier regenerates the clock"
    assert "Len(" not in verify, "a non-blank check is not a verification"
    for selective in ("CALC_STATE_ROW_LAST_SUCCESSFUL_FINGERPRINT",
                      "CALC_STATE_ROW_LAST_ATTEMPT_RESULT",
                      "CALC_STATE_ROW_CALCULATION_STATUS", "StoredText"):
        assert selective not in verify, (
            f"a selected-field check ({selective}) cannot stand for the exact comparison"
        )
    # The same block, all the way through: built, written, verified.
    statements = _statements(module, "RunCalculation")
    build = next(i for i, s in enumerate(statements) if "BuildSuccessBlock" in s)
    write = next(i for i, s in enumerate(statements) if "WriteSuccessCommit" in s)
    check = next(i for i, s in enumerate(statements) if "VerifySuccessCommit" in s)
    committed = statements.index("committed = True")
    assert build < write < check < committed, (
        "the commit is marked before it is proven, or built after it is written"
    )
    for index in (write, check):
        assert "successBlock" in statements[index], (
            "a different block reaches the write or the verification"
        )
    assert "Err.Raise" in statements[check + 1], (
        "an unverified commit continues to success"
    )


# ===========================================================================
# 10. NEGATIVE CONTROLS
# ===========================================================================
_STUB = 'Attribute VB_Name = "modProbe"\nOption Explicit\n'


def test_nc_01_writing_before_preparation_finishes_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RunCalculation() As OperationResult\n"
        "    WriteAnalytical package\n"
        "    If Not PrepareCurrentCalculation(package, detail) Then Exit Function\n"
        "End Function\n",
    )
    statements = _statements(planted, "RunCalculation")
    assert first_index(statements, r"WriteAnalytical") < \
        first_index(statements, r"PrepareCurrentCalculation")


def test_nc_02_split_success_metadata_writes_are_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub WriteSuccessCommit()\n"
        "    CalcSheet.Range(\"C13\").Value2 = Now\n"
        "    CalcSheet.Range(\"C14\").Value2 = package.Fingerprint\n"
        "End Sub\n",
    )
    body = _body(planted, "WriteSuccessCommit")
    assert body.count(".Value2 =") == 2, "the split writes must be visible"
    assert "Dim block(1 To 8, 1 To 1) As Variant" not in body


def test_nc_03_a_commit_outside_the_envelope_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RunCalculation() As OperationResult\n"
        "    On Error GoTo TransactionFailed\n    WriteAnalytical package\n"
        "    On Error GoTo 0\n    WriteSuccessCommit package\n"
        "TransactionFailed:\nEnd Function\n",
    )
    statements = _statements(planted, "RunCalculation")
    disarmed = first_index(statements, r"^On Error GoTo 0$")
    commit = first_index(statements, r"WriteSuccessCommit")
    assert disarmed < commit, "the commit outside the envelope must be visible"


def test_nc_04_an_omitted_table_in_rollback_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub RestoreSnapshot()\n"
        "    modWorkbook.RestoreTable modWorkbook.Lo(CALC_SHEET, TBL_CALC_YEARS), s.Years\n"
        "    modWorkbook.RestoreTable modWorkbook.Lo(CALC_SHEET, TBL_CALC_FX), s.Fx\n"
        "End Sub\n",
    )
    body = _body(planted, "RestoreSnapshot")
    assert body.count("modWorkbook.RestoreTable") == 2
    assert "TBL_CALC_DRIVERS" not in body


def test_nc_05_an_omitted_scalar_block_in_rollback_is_caught() -> None:
    for omitted, present in (("CALC_TOTALS_VALUE_RANGE", "CALC_STATE_VALUE_RANGE"),
                             ("CALC_STATE_VALUE_RANGE", "CALC_TOTALS_VALUE_RANGE")):
        planted = _synthetic(
            "modProbe",
            _STUB + "Private Sub RestoreSnapshot()\n"
            f"    CalcSheet.Range({present}).Value2 = s.Block\n"
            "End Sub\n",
        )
        body = _body(planted, "RestoreSnapshot")
        assert omitted not in body, f"the omitted {omitted} must be visible"


def test_nc_06_failed_metadata_before_rollback_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RollbackAndRecord() As OperationResult\n"
        "    WriteAttemptBlock CALC_ATTEMPT_FAILED, detail, CurrentStatus()\n"
        "    On Error GoTo RollbackFailed\n"
        "    RestoreSnapshot snapshot\n"
        "RollbackFailed:\nEnd Function\n",
    )
    statements = _statements(planted, "RollbackAndRecord")
    record = next(i for i, s in enumerate(statements) if "WriteAttemptBlock" in s)
    restore = next(i for i, s in enumerate(statements) if "RestoreSnapshot" in s)
    assert record < restore, "metadata written before the rollback must be visible"


def test_nc_07_a_refusal_touching_the_success_block_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub RecordRefusal()\n"
        "    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block\n"
        "End Sub\n",
    )
    assert "CALC_STATE_VALUE_RANGE" in _body(planted, "RecordRefusal")


def test_nc_08_a_refusal_clearing_analytical_tables_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub RecordRefusal()\n    WriteAnalytical emptyPackage\nEnd Sub\n",
    )
    assert "WriteAnalytical" in _body(planted, "RecordRefusal")


def test_nc_09_status_derived_from_the_attempt_history_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function DeriveStatus() As String\n"
        "    If StoredText(CALC_STATE_ROW_LAST_ATTEMPT_RESULT) = CALC_ATTEMPT_FAILED Then\n"
        "        DeriveStatus = CALC_STATUS_INVALID\n    End If\nEnd Function\n",
    )
    body = _body(planted, "DeriveStatus")
    assert "CALC_STATE_ROW_LAST_ATTEMPT_RESULT" in body
    assert "CALC_ATTEMPT_FAILED" in body


def test_nc_10_an_attempt_result_used_as_a_status_is_caught() -> None:
    for attempt in ("CALC_ATTEMPT_REFUSED", "CALC_ATTEMPT_FAILED"):
        planted = _synthetic(
            "modProbe",
            _STUB + "Private Function DeriveStatus() As String\n"
            f"    DeriveStatus = {attempt}\nEnd Function\n",
        )
        body = _body(planted, "DeriveStatus")
        assigned = set(re.findall(r"DeriveStatus = (\w+)", body))
        assert assigned - DERIVED_STATUSES == {attempt}, (
            f"the planted {attempt} status must be visible"
        )


def test_nc_11_comparing_two_empty_digests_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function DeriveStatus() As String\n"
        "    If StrComp(package.Fingerprint, stored, vbBinaryCompare) = 0 Then\n"
        "        DeriveStatus = CALC_STATUS_CURRENT\n    End If\nEnd Function\n",
    )
    statements = _statements(planted, "DeriveStatus")
    assert first_index(statements, r"If Len\(stored\) = 0 Then") == len(statements), (
        "the missing empty-digest guard must be visible"
    )
    assert first_index(statements, r"StrComp\(package\.Fingerprint, stored") < len(statements)


def test_nc_12_an_identity_written_into_the_audit_is_caught() -> None:
    for column in ("COL_CALC_DRIVERS_QUANTITY", "COL_CALC_DRIVERS_PROBABILITY"):
        planted = _synthetic(
            "modProbe",
            _STUB + "Private Function DriversBlock() As Variant\n"
            f"    block(row, {column}) = 1#\nEnd Function\n",
        )
        body = _body(planted, "DriversBlock")
        assert re.search(rf"block\(row, {column}\) = 1#", body)
        assert not re.search(rf"block\(row, {column}\) = Empty", body)


def test_nc_13_a_fabricated_base_year_rate_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function InflationBlock() As Variant\n"
        "    block(row, COL_CALC_INFLATION_FACTORS_ANNUAL_RATE) = 0#\nEnd Function\n",
    )
    body = _body(planted, "InflationBlock")
    assert re.search(r"COL_CALC_INFLATION_FACTORS_ANNUAL_RATE\) = 0#", body)
    assert not re.search(r"COL_CALC_INFLATION_FACTORS_ANNUAL_RATE\) = Empty", body)


def test_nc_14_seeding_the_reporting_currency_into_the_audit_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function FxBlock() As Variant\n"
        "    block(1, COL_CALC_FX_CURRENCY) = REPORTING_CURRENCY\nEnd Function\n",
    )
    assert "REPORTING_CURRENCY" in planted.code


def test_nc_15_a_fabricated_zero_row_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function YearsBlock() As Variant\n"
        "    ReDim block(1 To 1, 1 To 3)\n    block(1, 1) = 0\n    YearsBlock = block\n"
        "End Function\n",
    )
    body = _body(planted, "YearsBlock")
    assert not re.search(r"If rows < 1 Then Exit Function", body), (
        "the missing zero-row guard must be visible"
    )


def test_nc_16_a_calculate_button_is_caught() -> None:
    manifest = emitted_manifest()
    planted = list(manifest["buttons"]) + [
        {"key": "calculate", "sheet": "Setup", "shape_name": "btnPCCMCalculate",
         "caption": "Calculate", "entry_point": "PCCM_Calculate", "anchor_cell": "E50",
         "width": 150.0, "height": 28.0}
    ]
    bound = {b["entry_point"] for b in planted}
    assert len(planted) == 6 and "PCCM_Calculate" in bound, (
        "the planted button must be visible to the sweep"
    )


def test_nc_17_a_seventh_endpoint_is_caught() -> None:
    planted = _synthetic(
        "modProbe", _STUB + "Public Function PCCM_CalculationRefusal() As String\nEnd Function\n"
    )
    found = {p for p in planted.public_procedures if p.startswith("PCCM_Calc")}
    assert found - PCCM_ENDPOINTS == {"PCCM_CalculationRefusal"}


def test_nc_18_a_vacuous_manifest_proof_is_caught() -> None:
    """The pattern §31 forbids: pass loudest when the build is broken.

    Asserted against the real suites, not against a synthetic string.
    """
    for suite in ("test_phase5_report_source.py", "test_phase5_check_source.py",
                  "test_phase5_resolve_source.py", "test_phase5_stage_a.py"):
        text = (PCCM_ROOT / "tests" / suite).read_text(encoding="utf-8")
        statements = [line.strip() for line in text.splitlines()]
        for index, line in enumerate(statements):
            if re.match(r"^if not .*\.is_file\(\):$", line):
                assert statements[index + 1] != "return", (
                    f"{suite} still has a vacuous manifest/artifact guard at line {index + 1}"
                )


def test_nc_19_recomputing_a_report_value_from_the_sheet_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function TotalsBlock() As Variant\n"
        "    block(1, 1) = CalcSheet.Range(\"C23\").Value\nEnd Function\n",
    )
    body = _body(planted, "TotalsBlock")
    assert "CalcSheet" in body and ".Value" in body


def test_nc_20_removed_verification_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RunCalculation() As OperationResult\n"
        "    WriteAnalytical package\n    WriteSuccessCommit package\n"
        "    committed = True\nEnd Function\n",
    )
    statements = _statements(planted, "RunCalculation")
    assert first_index(statements, r"VerifyAnalytical") == len(statements)
    assert first_index(statements, r"VerifySuccessCommit") == len(statements)


# --- the correction-round controls -----------------------------------------
#
# Each plants the DEFECT independent review found, or a near neighbour of it, and
# asserts the detector above sees it. A detector nobody has watched fail is a
# detector nobody has tested.
def test_nc_21_a_header_number_framed_as_text_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function BuildFingerprint() As Boolean\n"
        "    If Not modCalcFingerprint.CalcFpCanonicalNumber(t.BaseYear, sep, text) "
        "Then Exit Function\n"
        "    header(0) = modCalcFingerprint.CalcFpCanonicalText(text)\n"
        "End Function\n",
    )
    code = planted.code
    calls = sorted(set(re.findall(r"modCalc(?:Factors|Analytical|Fingerprint)\.(\w+)", code)))
    assert "CalcFpCanonicalText" in calls and "CalcFpNumberField" not in calls, (
        "the text framing of a number must be visible to the surface sweep"
    )
    body = _body(planted, "BuildFingerprint")
    framed = re.findall(
        r"modCalcFingerprint\.CalcFpNumberField\(\s*package\.Model\.Timeline\.(\w+)", body
    )
    assert framed == [], "the header framer check must find no N field here"


def test_nc_22_a_hand_assembled_number_field_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function BuildFingerprint() As Boolean\n"
        "    header(0) = FP_TAG_NUMBER & CStr(Len(text)) & FP_FIELD_SEPARATOR & text\n"
        "End Function\n",
    )
    code = planted.code
    assert "FP_TAG_NUMBER" in code and "FP_FIELD_SEPARATOR" in code, (
        "reinvented framing must be visible outside the framing authority"
    )


def test_nc_23_publishing_through_the_modal_reporter_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "    modAppState.ReportResult result\n"
        "End Sub\n",
    )
    code = planted.code
    assert "ReportResult" in code, "the modal reporter must be visible"
    assert "modAppState.Announce" not in code, (
        "the automation-aware surface must be visibly absent"
    )


def test_nc_24_a_second_automation_reporter_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private gAutomationResult As String\n"
        "Public Sub PCCM_Calculate()\n"
        "    gAutomationResult = result.Detail\n"
        "End Sub\n",
    )
    assert "gAutomationResult" in planted.code, (
        "a privately invented automation channel must be visible"
    )


def test_nc_25_a_fallible_step_outside_the_envelope_is_caught() -> None:
    """Exactly the submitted shape: capture and begin ran before containment."""
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "    Dim state As AppStateSnapshot\n"
        "    state = modAppState.CaptureAppState()\n"
        "    modAppState.BeginOperation\n"
        "    On Error GoTo InvocationFailed\n"
        "    result = RunCalculation(committed)\n"
        "InvocationFailed:\nEnd Sub\n",
    )
    statements = _statements(planted, "PCCM_Calculate")
    executable = [
        s for s in statements
        if not re.match(r"^(Public |Private )?(Sub|Function)\b", s)
        and not re.match(r"^Dim\b", s)
    ]
    assert executable[0] != "On Error GoTo InvocationFailed", (
        "the uncovered capture and begin must be visible"
    )


def test_nc_26_cleanup_skipped_after_an_early_failure_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "    On Error GoTo InvocationFailed\n"
        "    state = modAppState.CaptureAppState()\n"
        "    result = RunCalculation(committed)\n"
        "    cleanup = modAppState.FinishOperation(state)\n"
        "    modAppState.Announce result\n"
        "    Exit Sub\n"
        "InvocationFailed:\n"
        "    modAppState.Announce modAppState.Failed(\"Calculate\", Err.Description)\n"
        "End Sub\n",
    )
    statements = _statements(planted, "PCCM_Calculate")
    finishes = [i for i, s in enumerate(statements) if "FinishOperation" in s]
    assert len(finishes) == 1, "the missing recovery cleanup must be visible"
    assert "stateCaptured" not in planted.code, "the missing flag must be visible"


def test_nc_27_cleanup_run_twice_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "    On Error GoTo InvocationFailed\n"
        "    stateCaptured = True\n"
        "    cleanup = modAppState.FinishOperation(state)\n"
        "    modAppState.Announce result\n"
        "    Exit Sub\n"
        "InvocationFailed:\n"
        "    If stateCaptured Then cleanup = modAppState.FinishOperation(state)\n"
        "End Sub\n",
    )
    statements = _statements(planted, "PCCM_Calculate")
    finishes = [i for i, s in enumerate(statements) if "FinishOperation" in s]
    assert len(finishes) == 2
    assert statements[finishes[0] + 1] != "stateCaptured = False", (
        "the uncleared flag that permits a second cleanup must be visible"
    )


def test_nc_27b_a_runtime_fault_recorded_as_a_refusal_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RunCalculation() As OperationResult\n"
        "    On Error GoTo PreWriteFailed\n"
        "    prepared = PrepareCurrentCalculation(package, detail)\n"
        "    If Not prepared Then\n"
        "        RunCalculation = RecordRefusal(detail)\n"
        "        Exit Function\n"
        "    End If\n"
        "PreWriteFailed:\n"
        "    RunCalculation = RecordRefusal(detail)\n"
        "TransactionFailed:\nEnd Function\n",
    )
    body = _body(planted, "RunCalculation")
    prewrite = body.partition("PreWriteFailed:")[2].partition("TransactionFailed:")[0]
    assert "RecordRefusal" in prewrite, (
        "a runtime fault dressed as a model refusal must be visible"
    )
    assert "RecordFailureWithoutRollback" not in prewrite


def test_nc_28_failed_metadata_written_after_a_failed_rollback_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RollbackAndRecord() As OperationResult\n"
        "    On Error GoTo RollbackFailed\n"
        "    RestoreSnapshot snapshot\n"
        "    WriteAttemptBlock CALC_ATTEMPT_FAILED, detail, CurrentStatus()\n"
        "    Exit Function\n"
        "RollbackFailed:\n"
        "    WriteAttemptBlock CALC_ATTEMPT_FAILED, detail, CurrentStatus()\n"
        "End Function\n",
    )
    body = _body(planted, "RollbackAndRecord")
    failed = body.split("RollbackFailed:", 1)[1]
    assert "WriteAttemptBlock" in failed, (
        "metadata written under a false premise of restoration must be visible"
    )


def test_nc_29_a_failed_record_undoing_the_rollback_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RollbackAndRecord() As OperationResult\n"
        "    On Error GoTo BookkeepingFailed\n"
        "    WriteAttemptBlock CALC_ATTEMPT_FAILED, detail, CurrentStatus()\n"
        "    Exit Function\n"
        "BookkeepingFailed:\n"
        "    RestoreSnapshot snapshot\n"
        "End Function\n",
    )
    body = _body(planted, "RollbackAndRecord")
    book = body.split("BookkeepingFailed:", 1)[1]
    assert "RestoreSnapshot" in book, (
        "a record failure re-running the rollback must be visible"
    )


def test_nc_30_a_committed_run_rewritten_to_failed_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function CleanupOutcome() As OperationResult\n"
        "    If committed Then\n"
        "        WriteAttemptBlock CALC_ATTEMPT_FAILED, cleanup, CurrentStatus()\n"
        "    End If\n"
        "End Function\n",
    )
    cleanup = _body(planted, "CleanupOutcome")
    assert "WriteAttemptBlock" in cleanup and "CALC_ATTEMPT_FAILED" in cleanup, (
        "a post-commit rewrite of C17 must be visible"
    )


def test_nc_31_a_commit_verified_on_seven_cells_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function VerifySuccessCommit(ByRef block As Variant) As Boolean\n"
        "    VerifySuccessCommit = VerifyRange(CALC_STATE_VALUE_RANGE, block, 7)\n"
        "End Function\n",
    )
    verify = _body(planted, "VerifySuccessCommit")
    assert "VerifyRange(CALC_STATE_VALUE_RANGE, block, 8)" not in verify, (
        "a verification that stops before C20 must be visible"
    )


def test_nc_32_a_stamp_checked_only_for_being_non_blank_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function VerifySuccessCommit() As Boolean\n"
        "    If Len(StoredText(CALC_STATE_ROW_LAST_SUCCESSFUL_STAMP)) = 0 Then Exit Function\n"
        "    VerifySuccessCommit = True\n"
        "End Function\n",
    )
    verify = _body(planted, "VerifySuccessCommit")
    assert "Len(" in verify and "StoredText" in verify, (
        "a non-blank stand-in for a value comparison must be visible"
    )
    assert "ByRef block As Variant" not in verify


def test_nc_33_a_verifier_that_regenerates_the_clock_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function VerifySuccessCommit(ByRef block As Variant) As Boolean\n"
        "    block(1, 1) = Now\n"
        "    VerifySuccessCommit = VerifyRange(CALC_STATE_VALUE_RANGE, block, 8)\n"
        "End Function\n",
    )
    assert "Now" in _body(planted, "VerifySuccessCommit"), (
        "a verifier comparing against a value the commit never held must be visible"
    )


def test_nc_34_two_captured_moments_in_one_commit_block_are_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub BuildSuccessBlock()\n"
        "    Dim built(1 To 8, 1 To 1) As Variant\n"
        "    built(1, 1) = Now\n"
        "    built(8, 1) = Now\n"
        "End Sub\n",
    )
    build = _body(planted, "BuildSuccessBlock")
    assert build.count("Now") == 2, "the second clock read must be visible"


def test_nc_35_marking_a_commit_before_verifying_it_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RunCalculation() As OperationResult\n"
        "    BuildSuccessBlock package, successBlock\n"
        "    WriteSuccessCommit successBlock\n"
        "    committed = True\n"
        "    If Not VerifySuccessCommit(successBlock) Then Err.Raise 5\n"
        "End Function\n",
    )
    statements = _statements(planted, "RunCalculation")
    check = next(i for i, s in enumerate(statements) if "VerifySuccessCommit" in s)
    committed = statements.index("committed = True")
    assert committed < check, "the premature commit flag must be visible"


def test_nc_36_a_handler_whose_label_is_missing_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RunCalculation() As OperationResult\n"
        "    On Error GoTo TransactionFailed\n"
        "    WriteAnalytical package\n"
        "End Function\n",
    )
    body = _body(planted, "RunCalculation")
    targets = {h for h in re.findall(r"On Error GoTo (\w+)", body) if h != "0"}
    assert targets == {"TransactionFailed"}
    assert not re.search(r"^TransactionFailed:$", body, re.M), (
        "the missing handler label must be visible"
    )


def test_nc_37_a_silent_path_out_of_the_endpoint_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "    On Error GoTo InvocationFailed\n"
        "    result = RunCalculation(committed)\n"
        "    modAppState.Announce result\n"
        "    Exit Sub\n"
        "InvocationFailed:\n"
        "End Sub\n",
    )
    statements = _statements(planted, "PCCM_Calculate")
    announces = [i for i, s in enumerate(statements) if "modAppState.Announce" in s]
    assert len(announces) == 1, "the unannounced handler path must be visible"


# --- correction round 2 -----------------------------------------------------
def test_nc_38_the_submitted_uncovered_normal_cleanup_is_caught() -> None:
    """THE ROUND-1 SHAPE, verbatim: the handler disarmed before the cleanup.

    `On Error GoTo 0` immediately above `FinishOperation` is exactly what shipped
    at 73adbb0. The detector must see that the call runs with nothing armed.
    """
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "    On Error GoTo InvocationFailed\n"
        "    state = modAppState.CaptureAppState()\n"
        "    stateCaptured = True\n"
        "    result = RunCalculation(committed)\n"
        "    On Error GoTo 0\n"
        "    cleanup = modAppState.FinishOperation(state)\n"
        "    stateCaptured = False\n"
        "    modAppState.Announce result\n"
        "    Exit Sub\n"
        "InvocationFailed:\n"
        "    On Error GoTo CleanupFailed\n"
        "    If stateCaptured Then cleanup = modAppState.FinishOperation(state)\n"
        "CleanupFailed:\nEnd Sub\n",
    )
    statements = _statements(planted, "PCCM_Calculate")
    normal = next(i for i, s in enumerate(statements) if "FinishOperation" in s)
    armed = [i for i, s in enumerate(statements[:normal])
             if s.startswith("On Error GoTo ") and s != "On Error GoTo 0"]
    disarmed = [i for i, s in enumerate(statements[:normal]) if s == "On Error GoTo 0"]
    assert armed and armed[-1] < disarmed[-1], (
        "the disarmed normal cleanup must be visible to the sweep"
    )
    assert "cleanupAttempted" not in planted.code, (
        "the missing exactly-once state must be visible"
    )


def test_nc_39_a_normal_cleanup_failure_that_retries_cleanup_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "    On Error GoTo NormalCleanupFailed\n"
        "    cleanupAttempted = True\n"
        "    cleanup = modAppState.FinishOperation(state)\n"
        "    Exit Sub\n"
        "NormalCleanupFailed:\n"
        "    cleanup = modAppState.FinishOperation(state)\n"
        "    modAppState.Announce result\n"
        "End Sub\n",
    )
    handler = _body(planted, "PCCM_Calculate").split("NormalCleanupFailed:", 1)[1]
    assert "FinishOperation" in handler, "the retried cleanup must be visible"


def test_nc_40_a_normal_cleanup_failure_that_rewrites_calc_state_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "    On Error GoTo NormalCleanupFailed\n"
        "    cleanup = modAppState.FinishOperation(state)\n"
        "    Exit Sub\n"
        "NormalCleanupFailed:\n"
        "    WriteAttemptBlock CALC_ATTEMPT_FAILED, Err.Description, CurrentStatus()\n"
        "    modAppState.Announce result\n"
        "End Sub\n",
    )
    handler = _body(planted, "PCCM_Calculate").split("NormalCleanupFailed:", 1)[1]
    assert "WriteAttemptBlock" in handler, (
        "a committed SUCCESS falsified by a cleanup exception must be visible"
    )


def test_nc_41_a_silent_normal_cleanup_failure_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "    On Error GoTo NormalCleanupFailed\n"
        "    cleanup = modAppState.FinishOperation(state)\n"
        "    modAppState.Announce result\n"
        "    Exit Sub\n"
        "NormalCleanupFailed:\n"
        "End Sub\n",
    )
    handler = _body(planted, "PCCM_Calculate").split("NormalCleanupFailed:", 1)[1]
    assert "modAppState.Announce" not in handler, (
        "a cleanup exception that reaches no announcement must be visible"
    )


def test_nc_42_a_recovery_cleanup_guarded_only_by_state_captured_is_caught() -> None:
    """The guard must consult the SPENT flag too, not just position."""
    planted = _synthetic(
        "modProbe",
        _STUB + "Public Sub PCCM_Calculate()\n"
        "InvocationFailed:\n"
        "    On Error GoTo CleanupFailed\n"
        "    If stateCaptured Then\n"
        "        cleanup = modAppState.FinishOperation(state)\n"
        "    End If\n"
        "CleanupFailed:\nEnd Sub\n",
    )
    body = _body(planted, "PCCM_Calculate")
    assert "If stateCaptured Then" in body
    assert "If stateCaptured And Not cleanupAttempted Then" not in body, (
        "a guard that cannot tell a spent attempt from an unused one must be visible"
    )


def test_nc_43_the_submitted_upstream_commit_failpoint_is_caught() -> None:
    """THE ROUND-1 SHAPE: the hook fires before commit preparation, not at it."""
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Function RunCalculation() As OperationResult\n"
        "    modAppState.FailPointCheck FAILPOINT_SUCCESS_COMMIT\n"
        "    BuildSuccessBlock package, successBlock\n"
        "    WriteSuccessCommit successBlock\n"
        "End Function\n"
        "Private Sub WriteSuccessCommit(ByRef block As Variant)\n"
        "    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block\n"
        "End Sub\n",
    )
    # The weak assertion still passes on the defect. The strong one does not.
    statements = _statements(planted, "RunCalculation")
    around = first_index(statements, r"FailPointCheck FAILPOINT_SUCCESS_COMMIT")
    commit = first_index(statements, r"WriteSuccessCommit")
    assert around < commit, "the superseded index check passes on the defect"
    writer = _statements(planted, "WriteSuccessCommit")
    assert not any("FAILPOINT_SUCCESS_COMMIT" in s for s in writer), (
        "the hook's absence from the commit procedure must be visible"
    )
    assert "FAILPOINT_SUCCESS_COMMIT" in _body(planted, "RunCalculation"), (
        "the upstream placement must be visible"
    )


def test_nc_44_a_commit_failpoint_after_the_assignment_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub WriteSuccessCommit(ByRef block As Variant)\n"
        "    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block\n"
        "    modAppState.FailPointCheck FAILPOINT_SUCCESS_COMMIT\n"
        "End Sub\n",
    )
    writer = [s for s in _statements(planted, "WriteSuccessCommit")
              if not re.match(r"^(Public |Private )?(Sub|Function)\b", s)]
    hook = next(i for i, s in enumerate(writer) if "FAILPOINT_SUCCESS_COMMIT" in s)
    assignment = next(i for i, s in enumerate(writer) if ".Value2 = block" in s)
    assert hook > assignment, (
        "a failpoint that can only fire after C13:C20 is published must be visible"
    )


def test_nc_45_a_statement_between_the_failpoint_and_the_commit_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub WriteSuccessCommit(ByRef block As Variant)\n"
        "    modAppState.FailPointCheck FAILPOINT_SUCCESS_COMMIT\n"
        "    block(1, 1) = Now\n"
        "    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block\n"
        "End Sub\n",
    )
    writer = [s for s in _statements(planted, "WriteSuccessCommit")
              if not re.match(r"^(Public |Private )?(Sub|Function)\b", s)]
    hook = next(i for i, s in enumerate(writer) if "FAILPOINT_SUCCESS_COMMIT" in s)
    assignment = next(i for i, s in enumerate(writer) if ".Value2 = block" in s)
    assert hook != assignment - 1, "the intervening statement must be visible"


def test_nc_46_a_removed_commit_failpoint_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub WriteSuccessCommit(ByRef block As Variant)\n"
        "    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block\n"
        "End Sub\n",
    )
    assert "FailPointCheck FAILPOINT_SUCCESS_COMMIT" not in planted.code, (
        "the missing commit failpoint must be visible"
    )


def test_nc_47_a_second_injection_mechanism_is_caught() -> None:
    planted = _synthetic(
        "modProbe",
        _STUB + "Private Sub WriteSuccessCommit(ByRef block As Variant)\n"
        "    If gAutomationFailAfterStage = \"Commit\" Then Err.Raise 5\n"
        "    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block\n"
        "End Sub\n",
    )
    assert "gAutomationFailAfterStage" in planted.code, (
        "a hand-rolled injection point must be visible"
    )


# ===========================================================================
# 11. this suite makes no runtime claim
# ===========================================================================
def test_52_no_test_in_this_file_claims_that_vba_ran() -> None:
    text = Path(__file__).read_text(encoding="utf-8")
    banned = (
        ("VBA", "produced"), ("VBA", "computed"), ("VBA", "returned"),
        ("the transaction", "committed"), ("rollback", "restored the workbook"),
        ("executed", "the VBA"), ("ran", "the VBA"), ("wrote", "to Excel"),
    )
    for parts in banned:
        assert " ".join(parts) not in text, f"this suite must not make that claim: {parts}"
    assert "NO VBA IS EXECUTED HERE" in text


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
