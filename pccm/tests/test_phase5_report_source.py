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
        "CalcFpBuildFingerprint", "CalcFpBuildRiskRecord", "CalcFpCanonicalNumber",
        "CalcFpCanonicalText", "Reconcile",
    ], f"unexpected numerical surface: {calls}"
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
    disarmed = first_index(statements, r"^On Error GoTo 0$")
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
    for guard in ("VerifyAnalytical(package)", "VerifySuccessCommit(package)"):
        index = next(i for i, t in enumerate(statements) if guard in t)
        assert any("Err.Raise" in t for t in statements[index:index + 3]), (
            f"a failed {guard} does not raise"
        )


def test_20_no_generic_error_suppression() -> None:
    code = _reporter().code
    assert "On Error Resume Next" not in code, (
        "a suppressed error would become a silently wrong calculation"
    )
    handlers = [h for h in re.findall(r"On Error GoTo (\w+)", code) if h != "0"]
    assert handlers == ["TransactionFailed"], (
        f"the only handler is the transaction envelope; found {handlers}"
    )


def test_21_both_failpoints_are_wired_through_the_phase_4_mechanism() -> None:
    """One where analytical state is half-written, one around the commit."""
    module = _reporter()
    statements = _statements(module, "RunCalculation")
    write = first_index(statements, r"WriteAnalytical")
    mid = first_index(statements, r"FailPointCheck FAILPOINT_ANALYTICAL_WRITE")
    commit = first_index(statements, r"WriteSuccessCommit")
    around = first_index(statements, r"FailPointCheck FAILPOINT_SUCCESS_COMMIT")
    assert write < mid, "the mid-write failpoint fires before anything is written"
    assert around < commit, "the commit failpoint fires after the commit"
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


def test_25_the_success_commit_is_one_eight_row_assignment() -> None:
    """Not four writes that could half-succeed and leave a fingerprint with no
    stamp, or a stamp with no version."""
    body = _body(_reporter(), "WriteSuccessCommit")
    assert "Dim block(1 To 8, 1 To 1) As Variant" in body
    assert body.count(".Value2 =") == 1
    assert "CALC_STATE_VALUE_RANGE" in body
    for row in range(1, 9):
        assert f"block({row}, 1) =" in body, f"commit row {row} is not populated"
    assert "FP_VERSION" in body, "C15 must carry the fingerprint version"
    assert "CALC_ATTEMPT_SUCCESS" in body and "CALC_STATUS_CURRENT" in body


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
    snapshot, exactly."""
    statements = _statements(_reporter(), "RunCalculation")
    restore = next(i for i, t in enumerate(statements) if "RestoreSnapshot" in t)
    record = next(i for i, t in enumerate(statements) if "RecordFailure" in t)
    assert restore < record, "FAILED metadata is written before the rollback"


def test_29_a_committed_operation_can_never_become_failed() -> None:
    statements = _statements(_reporter(), "RunCalculation")
    handler = statements.index("On Error GoTo 0", statements.index("committed = True"))
    tail = statements[handler:]
    guard = next(i for i, t in enumerate(tail) if t == "If committed Then")
    restore = next(i for i, t in enumerate(tail) if "RestoreSnapshot" in t)
    assert guard < restore, "a committed operation could still be rolled back"


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
    """FAILED is an attempt result. It never chooses the status."""
    module = _reporter()
    body = _body(module, "RecordFailure")
    assert "CurrentStatus()" in body, "the status is not re-derived after a rollback"
    assert "CALC_ATTEMPT_FAILED" in body
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
    body = _body(_reporter(), "PCCM_Calculate")
    for stage in ("modAppState.CaptureAppState()", "modAppState.BeginOperation",
                  "modAppState.FinishOperation(state)"):
        assert stage in body, f"{stage} is missing"
    statements = _statements(_reporter(), "PCCM_Calculate")
    run = next(i for i, t in enumerate(statements) if "RunCalculation()" in t)
    finish = next(i for i, t in enumerate(statements) if "FinishOperation" in t)
    assert run < finish, "application state is restored before the work is done"
    assert any("modAppState.Failed" in t for t in statements[finish:]), (
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
        "modCalcResolve": "3c67584390516a8a1c811df62d650749f6ef71518c649d7f1bb88dc753a837c1",
        "modCalcCheck": "738343945932150470233cb2a0b7e6fea7617db1a877cae8e09d19085e39c43b",
        "modCalcFactors": "721b8d6aa16fef850a13c714b329395730c9110ccd50d17c99927c3bfaae68c1",
        "modCalcAnalytical": "e234b3adacdb443c8c7b2b5072c311e7622405c3ec2e2987a750d85400299e0d",
        "modCalcFingerprint": "0a504c0dc29062420c5e4325117ef623157e5fc612de9a9e862d86315aed5802",
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
        _STUB + "Private Function RunCalculation() As OperationResult\n"
        "TransactionFailed:\n    RecordFailure detail\n    RestoreSnapshot snapshot\n"
        "End Function\n",
    )
    statements = _statements(planted, "RunCalculation")
    assert first_index(statements, r"RecordFailure") < \
        first_index(statements, r"RestoreSnapshot")


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
