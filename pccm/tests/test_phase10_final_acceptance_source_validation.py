#!/usr/bin/env python3
"""P10 FINAL WINDOWS ACCEPTANCE - mutation controls for the important requirements.

Each mutation breaks ONE thing the authorisation named, reruns the WHOLE
runner-source battery against the damaged copy, and requires a NAMED control
among the refusers. Nothing here writes to the repository; damaged copies live
in memory.

Runs standalone or under pytest.
"""
from __future__ import annotations

import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase10_final_acceptance_source as conformance  # noqa: E402


def _tests() -> list[str]:
    names = sorted(name for name in dir(conformance) if name.startswith("test_"))
    assert len(names) >= 30, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        if name in ("test_73_every_reachable_command_resolves_exactly_once",
                    "test_74_no_variable_can_be_read_before_assignment"):
            continue  # read the file on disk through PowerShell; not part of an in-memory mutation
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _mutate(expected: str, before: str, after: str) -> None:
    original = conformance._runner()
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:70]!r} is no longer in the runner")
    saved = dict(conformance._MEMO)
    conformance._MEMO["runner"] = damaged
    conformance._MEMO.pop("runner_code", None)
    try:
        refused = _run_battery()
    finally:
        conformance._MEMO.clear()
        conformance._MEMO.update(saved)
    assert refused, "the mutation survived the whole battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _mutate_source(expected: str, key: str, path: Path, before: str, after: str) -> None:
    """A production or runner text mutated IN MEMORY through the conformance
    module's memo, exactly as _mutate does for the runner."""
    original = conformance._src(key, path)
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(f"the mutation changed nothing: {before[:70]!r} is no longer in {path.name}")
    saved = dict(conformance._MEMO)
    conformance._MEMO[key] = damaged
    if key == "runner":
        conformance._MEMO.pop("runner_code", None)
    try:
        refused = _run_battery()
    finally:
        conformance._MEMO.clear()
        conformance._MEMO.update(saved)
    assert refused, "the mutation survived the whole battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def test_00_the_runner_as_written_passes_every_control() -> None:
    assert _run_battery() == []


def test_01_weakening_the_source_revision_comparison_is_refused() -> None:
    _mutate("test_22",
            "Add-FaCheck 'metadata.source-revision' ($observedRevision -ceq $revision.Expected)",
            "Add-FaCheck 'metadata.source-revision' ($observedRevision -ne '')")


def test_02_a_literal_module_count_is_refused() -> None:
    _mutate("test_30",
            "Add-FaCheck 'modules' ($moduleProblems.Count -eq 0)",
            "Add-FaCheck 'modules' ($observedModules.Count -eq 33)")


def test_03_a_performance_scale_request_is_refused() -> None:
    _mutate("test_06",
            "$acceptanceIterations = [int]$gateBCases.bounds.business_minimum_iterations",
            "$acceptanceIterations = 10000")


def test_04_dropping_the_reset_preservation_assertion_is_refused() -> None:
    _mutate("test_42",
            "    $null = Add-FaCheck 'reset.preserved' ($preservedAfter -ceq $preservedBefore) `\n",
            "    $null = Add-FaCheck 'reset.preserved' $true `\n")


def test_05_dropping_the_protection_assertion_after_a_refusal_is_refused() -> None:
    _mutate("test_62",
            "    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-refusal'\n",
            "")


def test_06_a_check_that_no_longer_fails_fast_is_refused() -> None:
    _mutate("test_64",
            "if ((-not $Ok) -and (-not $Continue)) {",
            "if ($false) {")


def test_07_reaching_for_excels_own_unprotect_is_refused() -> None:
    _mutate("test_12",
            "    $reply = [string]$Excel.Run('P10FW_Begin')\n",
            "    $reply = [string]$Excel.Run('P10FW_Begin')\n    $null = $Workbook.Worksheets.Item(1).Unprotect()\n")


def test_08_a_window_left_open_after_a_repair_precondition_is_refused() -> None:
    _mutate("test_05",
            "    try { Remove-TableRow -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $missingRowIndex }\n"
            "    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.missing-row' }\n",
            "    Remove-TableRow -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $missingRowIndex\n")


def test_09_ui_automation_is_refused() -> None:
    _mutate("test_13",
            "    $excel.Run('PCCM_AutomationEnd') | Out-Null\n",
            "    $excel.Run('PCCM_AutomationEnd') | Out-Null\n    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')\n")


def test_10_a_drifted_helper_copy_is_refused() -> None:
    _mutate("test_03",
            "if ($null -eq $Value) { $null = $rng.ClearContents() } else { $rng.Value2 = [double]$Value }",
            "$rng.Value2 = $Value")


def test_11_a_state_word_typed_as_a_literal_is_refused() -> None:
    _mutate("test_60",
            "$statusStale         = [string]$p7.model_states.derived_status[2]",
            "$statusStale         = 'STALE'")


def test_12_expecting_the_live_untouched_state_to_read_not_calculated_is_refused() -> None:
    """THE RUN-1 DEFECT, PUT BACK."""
    _mutate("test_60b",
            "(($states0.Calculation -ceq $statusInvalid) -and ($states0.Simulation -ceq $statusInvalid) -and",
            "(($states0.Calculation -ceq $statusNotCalculated) -and ($states0.Simulation -ceq $statusInvalid) -and")


def test_13_reading_the_persisted_history_after_the_compile_check_is_refused() -> None:
    """THE COMPILE CHECK EVALUATES AND PERSISTS THE LIVE STATUS; a persisted read
    after it would read INVALID and the assertion would be wrong for a second reason."""
    original = conformance._runner()
    start = original.index("    # 6a. THE PERSISTED CALCULATION HISTORY")
    end = original.index("    # 3. THE ACCEPTED COMPILE CHECK")
    block = original[start:end]
    moved = original[:start] + original[end:]
    anchor = "    # 2b. SHEETS AND CODENAMES"
    moved = moved.replace(anchor, block + anchor, 1)
    _mutate("test_60b", original, moved)


def test_14_expecting_pass_after_calculate_at_the_business_minimum_is_refused() -> None:
    """THE RUN-2 DEFECT, PUT BACK: no advisory expected, so the derived overall is PASS."""
    _mutate("test_60e",
            "-Scenario 'modelcheck.calculated' `\n        -Expected @($advisoryExpected)\n",
            "-Scenario 'modelcheck.calculated' `\n        -Expected @()\n")


def test_15_dropping_the_calculation_error_from_the_invalid_checkpoint_is_refused() -> None:
    _mutate("test_60f",
            "-Expected @(@{ Id = [string]$calcErrorChecks[0].check_id; Severity = $severityError; Subject = $victimId },\n"
            "                    $advisoryExpected, $annualHistoricalExpected, $annualHistoricalExpected)\n",
            "-Expected @($advisoryExpected, $annualHistoricalExpected, $annualHistoricalExpected)\n")


def test_16_dropping_the_not_calculated_warning_after_reset_is_refused() -> None:
    _mutate("test_60g",
            "-Scenario 'modelcheck.after-reset' `\n        -Expected @($advisoryExpected, $notCalculatedExpected)\n",
            "-Scenario 'modelcheck.after-reset' `\n        -Expected @($advisoryExpected)\n")


def test_17_tolerating_an_unrelated_actionable_row_is_refused() -> None:
    _mutate("test_60h",
            "    if ($unmatched.Count -gt 0) { $problems += ('unexpected actionable row(s): ' + (Format-FaActionable $unmatched)) }\n",
            "")


def test_18_reading_declared_checks_at_the_projection_root_is_refused() -> None:
    """THE ATTEMPT-3 DEFECT, PUT BACK."""
    _mutate("test_60k",
            "$calcErrorChecks = @($projection.evaluation.declared_checks | Where-Object {",
            "$calcErrorChecks = @($projection.declared_checks | Where-Object {")


def test_19_restoring_the_direct_optional_property_access_is_caught_by_the_executed_matcher() -> None:
    """THE RUN-4 DEFECT, PUT BACK, AND EXECUTED: the advisory case must die with
    PropertyNotFoundException under StrictMode exactly as Windows did."""
    import pytest as _pytest
    if not Path(conformance.PWSH).exists():
        _pytest.skip("no PowerShell on this host")
    original = conformance._runner()
    damaged = original.replace(
        "            if ($entry.HasSubject -and ((Format-FaCell $row.subject) -cne $entry.Subject)) { continue }\n",
        "            if (($null -ne $raw.Subject) -and ((Format-FaCell $row.subject) -cne [string]$raw.Subject)) { continue }\n", 1)
    assert damaged != original
    import tempfile
    scratch = Path(tempfile.mkdtemp(prefix="pccm-fa-matcher-")) / "phase10_final_acceptance.ps1"
    scratch.write_text(damaged, encoding="utf-8")
    lines = conformance._matcher_lines(scratch)
    assert lines["A.advisory"].startswith("ERROR|A.advisory|System.Management.Automation.PropertyNotFoundException|"), lines["A.advisory"]
    assert "'Subject'" in lines["A.advisory"]
    # And the static control refuses the same restoration without executing it.
    _mutate("test_60n",
            "            if ($entry.HasSubject -and ((Format-FaCell $row.subject) -cne $entry.Subject)) { continue }\n",
            "            if (($null -ne $raw.Subject) -and ((Format-FaCell $row.subject) -cne [string]$raw.Subject)) { continue }\n")


def test_20_dropping_a_malformed_definition_refusal_is_refused() -> None:
    _mutate("test_60n",
            "    if ($hasId -and $hasAnyOf) { throw ($where + 'names both Id and AnyOf; exactly one selector is allowed') }\n",
            "")


def test_21_omitting_one_historical_annual_warning_from_the_invalid_checkpoint_is_refused() -> None:
    """THE RUN-5 CLASS, PUT BACK BY HALF."""
    _mutate("test_60",
            "                    $advisoryExpected, $annualHistoricalExpected, $annualHistoricalExpected)\n",
            "                    $advisoryExpected, $annualHistoricalExpected)\n")


def test_22_omitting_both_historical_annual_warnings_is_refused() -> None:
    """THE RUN-5 DEFECT, PUT BACK."""
    _mutate("test_60",
            "                    $advisoryExpected, $annualHistoricalExpected, $annualHistoricalExpected)\n",
            "                    $advisoryExpected)\n")


def test_23_restoring_the_blank_token_for_the_historical_annual_subject_is_refused() -> None:
    """THE RUN-6 DEFECT, PUT BACK."""
    _mutate("test_60r",
            "$annualHistoricalExpected = @{ AnyOf = $annualWarningIds; Severity = $severityWarning; Subject = '' }",
            "$annualHistoricalExpected = @{ AnyOf = $annualWarningIds; Severity = $severityWarning; Subject = '<blank>' }")


def test_24_restoring_the_blank_token_for_the_not_calculated_subject_is_refused() -> None:
    _mutate("test_60r",
            "$notCalculatedExpected = @{ AnyOf = $calcWarningIds; Severity = $severityWarning; Subject = '' }",
            "$notCalculatedExpected = @{ AnyOf = $calcWarningIds; Severity = $severityWarning; Subject = '<blank>' }")


def test_25_dropping_the_subject_constraint_from_the_historical_annual_expectation_is_refused() -> None:
    """WITHOUT IT THE CONFIDENCE-LEVEL ANNUAL WARNING COULD PASS AS HISTORICAL."""
    _mutate("test_60r",
            "$annualHistoricalExpected = @{ AnyOf = $annualWarningIds; Severity = $severityWarning; Subject = '' }",
            "$annualHistoricalExpected = @{ AnyOf = $annualWarningIds; Severity = $severityWarning }")


def test_26_letting_one_row_satisfy_both_historical_expectations_is_refused() -> None:
    """THE MATCHED ROW MUST LEAVE THE POOL - proved on the runner's text and by
    executing the damaged matcher: with the removal gone, the second historical
    entry re-matches the first row and both rows are then reported unexpected."""
    import pytest as _pytest
    removal = "            $unmatched = @($unmatched | Where-Object { -not [object]::ReferenceEquals($_, $found) })\n"
    original = conformance._runner()
    assert original.count(removal) == 1
    if Path(conformance.PWSH).exists():
        import tempfile
        scratch = Path(tempfile.mkdtemp(prefix="pccm-fa-distinct-")) / "phase10_final_acceptance.ps1"
        scratch.write_text(original.replace(removal, "", 1), encoding="utf-8")
        lines = conformance._matcher_lines(scratch)
        assert lines["K.distinct-rows"].startswith("MATCH|K.distinct-rows|ok=False|"), lines["K.distinct-rows"]
        assert lines["G.invalid-after-annual"].startswith("MATCH|G.invalid-after-annual|ok=False|"), lines["G.invalid-after-annual"]
    _mutate("test_60h", removal, "")


# ---------------------------------------------------------------------------
# THE BOUNDED CORRECTION ROUND AFTER THE INDEPENDENT REVIEW
# ---------------------------------------------------------------------------
def test_27_accepting_a_refused_shrink_as_success_is_refused() -> None:
    """THE GRID COMPARISON AROUND A REFUSAL REQUIRES A FAIL ANNOUNCEMENT."""
    _mutate("test_53",
            "        $null = Add-FaCheck $Scenario (($Announcement -like 'FAIL|*') -and ($costNow -ceq $script:FaCostBefore) -and ($riskNow -ceq $script:FaRiskBefore)) `\n",
            "        $null = Add-FaCheck $Scenario (($Announcement -like '*|*') -and ($costNow -ceq $script:FaCostBefore) -and ($riskNow -ceq $script:FaRiskBefore)) `\n")


def test_28_expecting_the_typed_zero_shrink_to_succeed_is_refused() -> None:
    _mutate("test_53",
            "    Assert-FaGridsUnchanged -Scenario 'repair.shrink-zero-refused' -Announcement $zeroRefused\n",
            "    $null = Add-FaCheck 'repair.shrink-zero-refused' ($zeroRefused -like 'OK|*') $zeroRefused\n")


def test_29_reading_a_populated_zero_total_as_blank_is_refused() -> None:
    """THE SIGNED ROW MUST BE REFUSED AS POPULATED, NOT ALLOWED AS EMPTY."""
    _mutate("test_53",
            "    Assert-FaGridsUnchanged -Scenario 'repair.signed-zero-total-refused' -Announcement $signedRefused\n",
            "    $null = Add-FaCheck 'repair.signed-zero-total-refused' ($signedRefused -like 'OK|*') $signedRefused\n")


def test_30_a_regrown_project_year_expected_to_hold_zero_is_refused() -> None:
    _mutate("test_52",
            "        if ([string]$after[$before.Count - 1] -ne '') { $growthProblems += ([string]$before[0] + ' regrown project year '",
            "        if ([string]$after[$before.Count - 1] -ne '0') { $growthProblems += ([string]$before[0] + ' regrown project year '")


def test_31_a_column_added_outside_the_window_is_refused() -> None:
    _mutate("test_51",
            "    try { $null = Add-FaTableColumn -Workbook $wb -SheetName $gridSheet -TableName $gridTable }\n"
            "    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-blank' }\n",
            "    $null = Add-FaTableColumn -Workbook $wb -SheetName $gridSheet -TableName $gridTable\n")


def test_32_dropping_the_rollback_digest_comparison_is_refused() -> None:
    _mutate("test_54",
            "             ($costNow -ceq $script:FaCostBefore) -and ($riskNow -ceq $script:FaRiskBefore)) `\n",
            "             $true) `\n")


def test_33_a_com_exception_counted_as_protection_is_refused() -> None:
    """THE OLD DEFECT, RE-INTRODUCED: an exception on the write path sets the
    check's success. The restated control refuses it."""
    _mutate("test_63",
            "    } catch { $codeWriteFailure = (Format-Err $_) }\n",
            "    } catch { $codeWriteFailure = '' }\n")


def test_34_expecting_the_locked_cell_write_to_fail_is_refused() -> None:
    """THE INDEPENDENT REVIEW'S FINDING: UserInterfaceOnly permits the write."""
    _mutate("test_63",
            "    $null = Add-FaCheck 'protection.locked-cell.code-write-permitted' (($codeWriteFailure -eq '') -and ($valueAfterWrite -ceq [string]$sourceRevisionRow.label)) `\n",
            "    $null = Add-FaCheck 'protection.locked-cell.code-write-permitted' ($codeWriteFailure -ne '') `\n")


def test_35_dropping_the_user_edit_protection_half_is_refused() -> None:
    _mutate("test_63",
            "    $null = Add-FaCheck 'protection.locked-cell.user-protected' (($lockedReadFailure -eq '') -and $lockedIsLocked -and $sheetIsProtected) `\n",
            "    $null = Add-FaCheck 'protection.locked-cell.user-protected' ($lockedReadFailure -eq '') `\n")


def test_36_a_saved_distribution_copy_is_refused() -> None:
    _mutate("test_55",
            "    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'copy.protection-after-run'\n",
            "    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'copy.protection-after-run'\n    $wb.Save()\n")


def test_37_a_copy_opened_from_the_original_path_is_refused() -> None:
    _mutate("test_55",
            "    Copy-Item -LiteralPath $stageBPath -Destination $copyPath -Force\n    $wb = $workbooks.Open($copyPath)\n",
            "    $wb = $workbooks.Open($stageBPath)\n")


def test_38_an_undeclared_edit_to_the_executed_tail_is_refused() -> None:
    """THE REVERSAL: any byte moved outside the two delimited insertions and
    the protection substitution breaks the identity with the executed runner."""
    _mutate("test_60c",
            "    $null = Add-FaCheck 'repair.fingerprint' (($recalc2 -like 'OK|*') -and ($fingerprintAfterRepairs -ceq $fingerprint)) `\n",
            "    $null = Add-FaCheck 'repair.fingerprint' ($recalc2 -like 'OK|*') `\n")


# ---------------------------------------------------------------------------
# THE P10-R3 CLOSURE: CONTRACT MATRIX ROW O
# ---------------------------------------------------------------------------
def test_39_leaving_the_event_switch_off_is_refused() -> None:
    _mutate("test_56",
            "    } finally { $excel.EnableEvents = $true }\n",
            "    } finally { $null = $true }\n")


def test_40_arming_with_the_event_switch_on_is_refused() -> None:
    """THE HANDLER WOULD RUN AT OPEN, unarmed, and the row would prove nothing."""
    _mutate("test_56",
            "    $excel.EnableEvents = $false\n    try {\n        $wb = $workbooks.Open($openPath)\n",
            "    try {\n        $wb = $workbooks.Open($openPath)\n")


def test_41_dropping_the_real_handler_run_is_refused() -> None:
    _mutate("test_56",
            "    try { $excel.Run(\"'\" + [string]$wb.Name + \"'!ThisWorkbook.Workbook_Open\") | Out-Null }\n    catch { $handlerFailure = (Format-Err $_) }\n",
            "    $handlerFailure = ''\n")


def test_42_an_exception_from_the_run_counted_as_the_failure_path_is_refused() -> None:
    _mutate("test_56",
            "    catch { $handlerFailure = (Format-Err $_) }\n",
            "    catch { $handlerFailure = '' }\n")


def test_43_a_weakened_record_expectation_is_refused() -> None:
    _mutate("test_56",
            "    $null = Add-FaCheck 'rowo.failure-path' ($recorded -ceq $expectedRecord) `\n",
            "    $null = Add-FaCheck 'rowo.failure-path' ($recorded -clike 'Workbook_Open: *') `\n")


def test_44_the_runner_restoring_screen_updating_itself_is_refused() -> None:
    """THE HANDLER MUST RESTORE IT; a runner write would hide a handler that did not."""
    _mutate("test_56",
            "    $recorded = Get-FaRunText -Excel $excel -Procedure 'PCCM_AutomationResult'\n    $prompted",
            "    $excel.ScreenUpdating = $screenBefore\n    $recorded = Get-FaRunText -Excel $excel -Procedure 'PCCM_AutomationResult'\n    $prompted")


def test_45_tolerating_a_half_protected_workbook_is_refused() -> None:
    _mutate("test_56",
            "((-not $released.Applied) -and ($released.Protected -eq 0) -and (-not $released.Structure) -and ($released.Depth -eq 0))",
            "((-not $released.Applied) -and ($released.Protected -ge 0) -and (-not $released.Structure) -and ($released.Depth -eq 0))")


def test_46_a_failpoint_name_that_is_not_the_handlers_is_refused() -> None:
    _mutate("test_56",
            "$script:OpenFailpoint = 'Phase10WorkbookOpen'\n",
            "$script:OpenFailpoint = 'Phase10OpenSomethingElse'\n")


def test_47_an_undeclared_edit_outside_the_row_o_block_is_refused() -> None:
    _mutate("test_60c",
            "    $null = Add-FaCheck 'copy.fixture' ($copyApplied -like 'OK|*') $copyApplied\n",
            "    $null = Add-FaCheck 'copy.fixture' $true $copyApplied\n")


def test_48_a_failpoint_before_the_apply_is_refused() -> None:
    """IT WOULD BYPASS THE RELEASE: nothing applied, nothing to put right."""
    _mutate_source("test_57", "handler", conformance.HANDLER_VBA,
                   "    If Not modProtection.ProtectionApply(detail) Then GoTo Failed\n    modAppState.FailPointCheck FAILPOINT_WORKBOOK_OPEN\n",
                   "    modAppState.FailPointCheck FAILPOINT_WORKBOOK_OPEN\n    If Not modProtection.ProtectionApply(detail) Then GoTo Failed\n")


def test_49_removing_the_failpoint_is_refused() -> None:
    _mutate_source("test_57", "handler", conformance.HANDLER_VBA,
                   "    modAppState.FailPointCheck FAILPOINT_WORKBOOK_OPEN\n", "")


def test_50_reading_err_after_the_release_is_refused() -> None:
    """THE RELEASE OWNER'S On Error CLEARS Err; read afterwards, the record names nothing."""
    _mutate_source("test_57", "handler", conformance.HANDLER_VBA,
                   "    If Len(detail) = 0 Then detail = Err.Description\n    ' A HALF-PROTECTED WORKBOOK IS THE ONE OUTCOME TO AVOID.",
                   "    ' A HALF-PROTECTED WORKBOOK IS THE ONE OUTCOME TO AVOID.")


def test_51_production_beginning_the_seam_is_refused() -> None:
    """THE TRIGGER MUST STAY DORMANT: no production line may arm automation."""
    _mutate_source("test_57", "handler", conformance.HANDLER_VBA,
                   "    restored = False\n",
                   "    restored = False\n    modAppState.gAutomationActive = True\n")


def test_52_a_failpoint_that_fires_without_the_seam_is_refused() -> None:
    _mutate_source("test_57", "appstate", conformance.APPSTATE_VBA,
                   "    If Not gAutomationActive Then Exit Sub\n    If Len(gAutomationFailAfterStage) = 0 Then Exit Sub\n",
                   "    If Len(gAutomationFailAfterStage) = 0 Then Exit Sub\n")


def test_53_a_second_statement_on_the_successful_path_is_refused() -> None:
    """NORMAL Workbook_Open BEHAVIOUR IS PINNED statement by statement."""
    _mutate_source("test_57", "handler", conformance.HANDLER_VBA,
                   "    Application.ScreenUpdating = previousUpdating\n    restored = True\n",
                   "    Application.ScreenUpdating = previousUpdating\n    Application.Calculation = xlCalculationAutomatic\n    restored = True\n")


# ---------------------------------------------------------------------------
# FINAL ACCEPTANCE RUN 8: THE WIDTH-GROWTH FIXTURE
# ---------------------------------------------------------------------------
def test_54_shaping_only_row_one_is_refused() -> None:
    """RUN 8's DEFECT, RE-INTRODUCED: another populated row is left non-100%."""
    _mutate("test_52",
            "    for ($r = 0; $r -lt $originalCostBody.Count; $r++) {\n        if ([string]$originalCostBody[$r][0] -eq '') { continue }\n",
            "    for ($r = 0; $r -lt 1; $r++) {\n        if ([string]$originalCostBody[$r][0] -eq '') { continue }\n")


def test_55_skipping_the_fixture_restoration_is_refused() -> None:
    _mutate("test_52",
            "    Restore-FaCostRows -Original $originalCostBody\n",
            "")


def test_56_restoring_a_blank_as_a_zero_is_refused() -> None:
    _mutate("test_52",
            "-RowIndex $index -Year $y -Weight $null }",
            "-RowIndex $index -Year $y -Weight 0.0 }")


def test_57_dropping_the_fixture_precondition_assertion_is_refused() -> None:
    _mutate("test_52",
            "    $null = Add-FaCheck 'repair.width-growth.fixture' ($fixtureProblems.Count -eq 0) `\n",
            "    $null = Add-FaCheck 'repair.width-growth.fixture' $true `\n")


def test_58_not_assessing_the_risk_grid_is_refused() -> None:
    _mutate("test_52",
            "    $fixtureProblems += @(Get-FaProfileProblems -Body $riskBeforeGrowth -FixedColumns $riskFixedColumns -YearCount $durationYears -Label 'Risk Profiling')\n",
            "")


def test_59_a_relaxed_restoration_proof_is_refused() -> None:
    _mutate("test_52",
            "    $null = Add-FaCheck 'repair.width-growth.restored' (($costAfterGrowthRestore -ceq $baselineCost) -and ($riskAfterGrowthRestore -ceq $baselineRisk)) `\n",
            "    $null = Add-FaCheck 'repair.width-growth.restored' ($costAfterGrowthRestore -ceq $baselineCost) `\n")


def test_60_weakening_the_production_semantic_gate_for_the_harness_is_refused() -> None:
    """PRODUCTION WAS RIGHT AT RUN 8 and is pinned byte for byte to the candidate."""
    _mutate_source("test_58", "repair", conformance.REPAIR_VBA,
                   "    If WithinTolerance(total, REPAIR_PROFILE_SUM_TARGET) Then\n",
                   "    If WithinTolerance(total, REPAIR_PROFILE_SUM_TARGET) Or (total = 0.75) Then\n")


def test_61_an_undeclared_runner_edit_outside_the_width_growth_scenario_is_refused() -> None:
    _mutate("test_60c4",
            "    $null = Add-FaCheck 'repair.shrink-blank' (($shrunkBlank -like 'OK|*') -and ($costAfterShrink -ceq $script:FaCostBefore)) `\n",
            "    $null = Add-FaCheck 'repair.shrink-blank' ($shrunkBlank -like 'OK|*') `\n")


if __name__ == "__main__":
    failures = 0
    for name in sorted(n for n in dir() if n.startswith("test_")):
        try:
            globals()[name]()
            print(f"PASS {name}")
        except BaseException as exc:  # noqa: BLE001
            failures += 1
            print(f"FAIL {name}: {exc}")
    sys.exit(1 if failures else 0)
