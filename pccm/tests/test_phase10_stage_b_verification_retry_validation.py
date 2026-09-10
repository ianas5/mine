#!/usr/bin/env python3
"""P10 / STAGE-B RETRY MUTATION CONTROLS.

A RETRY HELPER IS THE EASIEST THING IN THIS REPOSITORY TO MAKE USELESS. Widen
the HRESULT set to "COMException", drop one bound, or return a value instead of
rethrowing, and the bootstrap starts reporting success it did not observe --
which is the one outcome Stage-B verification exists to prevent.

So every control in the conformance suite is fed the exact damage it claims to
refuse. Each mutation breaks ONE thing, reruns the WHOLE battery against the
damaged copy, and requires a NAMED control among the refusers. A mutation that
survives means the control it belongs to is decoration.

Nothing here writes to the repository. Damaged copies live in memory.

Runs standalone or under pytest.
"""
from __future__ import annotations

import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase10_stage_b_verification_retry as conformance  # noqa: E402


def _tests() -> list[str]:
    names = sorted(name for name in dir(conformance) if name.startswith("test_"))
    assert len(names) >= 30, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _mutate(key: str, expected: str, before: str, after: str) -> None:
    """Damage one source, and prove the anchor still matched.

    NOT AN AssertionError WHEN THE ANCHOR MISSES. A no-op mutation that raised one
    would be indistinguishable from the refusal it is meant to provoke, and the
    suite would report a clean sweep while testing nothing.
    """
    loader = {"lifecycle": conformance._lifecycle, "build": conformance._build}[key]
    original = loader()
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:70]!r} is no longer in {key}")
    saved = dict(conformance._MEMO)
    conformance._MEMO[key] = damaged
    conformance._MEMO.pop(f"{key}_code", None)
    try:
        refused = _run_battery()
    finally:
        conformance._MEMO.clear()
        conformance._MEMO.update(saved)
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


# ===========================================================================
# A. THE RETRY SET
# ===========================================================================
def test_01_retrying_every_comexception_is_refused() -> None:
    """THE NAMED MUTATION. Retrying on the exception TYPE rather than the refusal
    code turns 'the call never ran' into 'something went wrong, try again'."""
    _mutate("lifecycle", "test_04",
            "            $name = Get-ComRejectionName $_\n",
            "            $name = ''\n"
            "            if ($_.Exception -is [System.Runtime.InteropServices.COMException]) "
            "{ $name = 'COMException' }\n")


def test_02_adding_a_third_hresult_to_the_table_is_refused() -> None:
    _mutate("lifecycle", "test_01",
            "    -2147417846 = 'RPC_E_SERVERCALL_RETRYLATER (0x8001010A)'\n",
            "    -2147417846 = 'RPC_E_SERVERCALL_RETRYLATER (0x8001010A)'\n"
            "    -2146777998 = 'VBA_E_IGNORE (0x800AC472)'\n")


def test_03_a_plausible_name_over_the_wrong_code_is_refused() -> None:
    """THE PROBE RUN 2 LESSON. A wrong value behind a right-looking name is the
    defect that cost two Windows runs, so the name carries its own hex."""
    _mutate("lifecycle", "test_02",
            "    -2147418111 = 'RPC_E_CALL_REJECTED (0x80010001)'",
            "    -2147418110 = 'RPC_E_CALL_REJECTED (0x80010001)'")


def test_04_silently_dropping_the_excluded_code_from_the_record_is_refused() -> None:
    """AN OMISSION IS NOT A DECISION."""
    _mutate("lifecycle", "test_03",
            "# VBA_E_IGNORE (0x800AC472) is deliberately NOT in the set.",
            "# The set is what it is.")


def test_05_classifying_inside_the_retry_loop_is_refused() -> None:
    _mutate("lifecycle", "test_04",
            "            $name = Get-ComRejectionName $_",
            "            $name = ''\n"
            "            if ([int]$_.Exception.ErrorCode -eq -2147418111) { $name = 'rejected' }")


# ===========================================================================
# B. THE BOUNDS
# ===========================================================================
def test_10_removing_the_attempt_limit_is_refused() -> None:
    """THE NAMED MUTATION: bound removal."""
    _mutate("lifecycle", "test_11",
            "            if ($attempts -ge $MaxAttempts) {",
            "            if ($false) {")


def test_11_removing_the_total_budget_check_is_refused() -> None:
    _mutate("lifecycle", "test_12",
            "            } elseif (($waitedMs + $delay) -gt $TotalBudgetMs) {",
            "            } elseif ($false) {")


def test_12_checking_the_budget_after_sleeping_is_refused() -> None:
    """A BUDGET CHECKED TOO LATE HAS ALREADY BEEN OVERSPENT."""
    # A GENUINE REORDERING: the sleep is hoisted above the exhaustion check, so
    # every refused call costs its delay whether or not the budget is spent.
    _mutate("lifecycle", "test_12",
            "            $reason = ''\n",
            "            Start-Sleep -Milliseconds $delay\n            $reason = ''\n")


def test_12a_renaming_the_budget_parameter_is_refused() -> None:
    """THE BOUND STILL HAS TO BE THE ONE WITH A DEFAULT."""
    _mutate("lifecycle", "test_10",
            "        [int]$TotalBudgetMs = 15000",
            "        [int]$Unbudgeted = 15000")


def test_13_an_infinite_default_budget_is_refused() -> None:
    _mutate("lifecycle", "test_10",
            "        [int]$TotalBudgetMs = 15000",
            "        [int]$TotalBudgetMs = 3600000")


def test_14_an_unbounded_attempt_default_is_refused() -> None:
    _mutate("lifecycle", "test_10",
            "        [int]$MaxAttempts   = 12,",
            "        [int]$MaxAttempts   = 100000,")


def test_15_letting_the_delay_grow_without_a_cap_is_refused() -> None:
    _mutate("lifecycle", "test_13",
            "            $delay = [Math]::Min(($delay + $FirstDelayMs), $MaxDelayMs)",
            "            $delay = $delay * 2")


def test_16_a_waited_counter_that_never_moves_is_refused() -> None:
    _mutate("lifecycle", "test_15",
            "            $waitedMs = $waitedMs + $delay",
            "            $waitedMs = $waitedMs")


def test_16a_moving_the_bound_out_of_the_loop_head_is_refused() -> None:
    """THE FIRST DRAFT OF THIS HELPER WROTE `while ($true)`, and the Phase-5
    bounded-loop control caught it. Both controls now refuse it."""
    _mutate("lifecycle", "test_14a",
            "    while ($attempts -lt $MaxAttempts) {",
            "    while ($true) {")


def test_16b_returning_a_record_after_falling_out_of_the_loop_is_refused() -> None:
    _mutate("lifecycle", "test_14b",
            "    if (-not $answered) {\n"
            "        throw ('Invoke-ComRetryRead: ' + $Description + ' was never answered after ' +\n"
            "               [string]$attempts + ' attempt(s).')\n"
            "    }\n",
            "")


def test_17_a_blind_sleep_before_verification_is_refused() -> None:
    """THE FIX THE AUTHORISATION FORBIDS STARTING WITH."""
    _mutate("build", "test_14",
            "        $wb2 = $workbooks2.Open($stageBPath)",
            "        $wb2 = $workbooks2.Open($stageBPath)\n"
            "        Start-Sleep -Seconds 10")


# ===========================================================================
# C. WHAT THE HELPER MUST NOT BECOME
# ===========================================================================
def test_20_a_scriptblock_parameter_is_refused() -> None:
    """IT WOULD MAKE EVERY CALL RETRYABLE, WRITES INCLUDED."""
    _mutate("lifecycle", "test_2",
            "        $Target,\n        [string]$Member,",
            "        $Target,\n        [scriptblock]$Action,\n        [string]$Member,")


def test_21_invoking_an_arbitrary_member_is_refused() -> None:
    _mutate("lifecycle", "test_21",
            "            if ($useItem) { $value = $Target.Item($Key) } else { $value = $Target.$Member }",
            "            if ($useItem) { $value = $Target.$Member($Key) } else { $value = $Target.$Member }")


def test_22_accepting_a_key_for_any_member_is_refused() -> None:
    _mutate("lifecycle", "test_22",
            "    if ($useItem -and $Member -ne 'Item') {",
            "    if ($false) {")


def test_23_returning_the_bare_value_is_refused() -> None:
    """THE RUN-3 COLLAPSE, ONE LAYER OUT: PowerShell would enumerate a returned
    Worksheets collection into its members."""
    _mutate("lifecycle", "test_23",
            "    return [pscustomobject]@{\n        Description = $Description\n        Value       = $value",
            "    return $value\n    $unreachable = [pscustomobject]@{\n        Description = $Description\n        Value       = $value")


def test_24_emitting_the_read_to_the_pipeline_is_refused() -> None:
    _mutate("lifecycle", "test_23",
            "            if ($useItem) { $value = $Target.Item($Key) } else { $value = $Target.$Member }",
            "            if ($useItem) { $Target.Item($Key) } else { $Target.$Member }")


def test_25_an_anonymous_retry_is_refused() -> None:
    """REQUIRED: the exact action being retried must be named in diagnostics."""
    _mutate("lifecycle", "test_24",
            "    if ([string]::IsNullOrWhiteSpace($Description)) {",
            "    if ($false) {")


def test_26_a_call_site_without_a_description_is_refused() -> None:
    _mutate("build", "test_24",
            "        $ff = [int](Invoke-ComRetryRead -Target $wb2 -Member 'FileFormat' `\n"
            "                        -Description 'the reopened workbook FileFormat').Value",
            "        $ff = [int](Invoke-ComRetryRead -Target $wb2 -Member 'FileFormat').Value")


def test_27_an_unbounded_inner_exception_walk_is_refused() -> None:
    _mutate("lifecycle", "test_25",
            "    for ($depth = 0; $depth -lt 5; $depth++) {",
            "    while ($true) {")


def test_28_a_classifier_that_falls_through_to_retry_is_refused() -> None:
    _mutate("lifecycle", "test_26",
            "        $ex = $next\n    }\n    return ''",
            "        $ex = $next\n    }\n    return 'unknown'")


# ===========================================================================
# D. FABRICATED SUCCESS
# ===========================================================================
def test_30_returning_a_default_after_exhaustion_is_refused() -> None:
    """'COULD NOT READ DASHBOARD, ASSUME OKAY' - named in the authorisation as
    the thing that must have no path."""
    _mutate("lifecycle", "test_06",
            "                $script:comRetryWaitMs = $script:comRetryWaitMs + $waitedMs\n                throw",
            "                $script:comRetryWaitMs = $script:comRetryWaitMs + $waitedMs\n"
            "                return [pscustomobject]@{ Description = $Description; Value = $null; "
            "Attempts = $attempts; WaitedMs = $waitedMs; Rejections = 'assumed okay' }")


def test_31_swallowing_the_original_error_is_refused() -> None:
    _mutate("lifecycle", "test_05",
            "            if ([string]::IsNullOrWhiteSpace($name)) {\n"
            "                # Not a refused call. Excel accepted this one and it failed, so it\n"
            "                # is a real result and it is rethrown untouched.\n"
            "                throw\n"
            "            }",
            "            if ([string]::IsNullOrWhiteSpace($name)) {\n"
            "                throw New-Object System.Exception ('read failed: ' + $Description)\n"
            "            }")


def test_32_dropping_the_per_sheet_problem_record_is_refused() -> None:
    _mutate("build", "test_41",
            '                $problems += ("{0}: {1}" -f $sheet.name, (Format-Err $_))',
            '                Write-Host "sheet unreadable, assuming okay"')


def test_33_storing_the_com_value_in_the_ledger_is_refused() -> None:
    """THE LIFECYCLE POLICY: a diagnostic collection never holds an Excel RCW."""
    _mutate("lifecycle", "test_53",
            "            ('RETRIED    ' + $Description + ' :: answered on attempt ' + [string]$attempts +",
            "            ('RETRIED    ' + [string]$value + $Description + ' :: answered on attempt ' + [string]$attempts +")


# ===========================================================================
# E. WHERE IT IS APPLIED
# ===========================================================================
def test_40_un_retrying_the_read_that_actually_failed_is_refused() -> None:
    """THE DASHBOARD READ ITSELF, put back the way it was when it failed."""
    _mutate("build", "test_3",
            "                $ws = (Invoke-ComRetryRead -Target $worksheets2 -Member 'Item' -Key ([string]$sheet.name) `\n"
            "                           -Description (\"Worksheets.Item('\" + [string]$sheet.name + \"')\")).Value",
            "                $ws = $worksheets2.Item($sheet.name)")


def test_41_un_retrying_the_codename_read_is_refused() -> None:
    _mutate("build", "test_3",
            "                $codeName = [string](Invoke-ComRetryRead -Target $ws -Member 'CodeName' `\n"
            "                                         -Description ([string]$sheet.name + '.CodeName')).Value",
            "                $codeName = [string]$ws.CodeName")


def test_42_un_retrying_a_button_read_is_refused() -> None:
    _mutate("build", "test_3",
            "                $onAction = [string](Invoke-ComRetryRead -Target $shp -Member 'OnAction' `\n"
            "                                         -Description ([string]$button.shape_name + '.OnAction')).Value",
            "                $onAction = [string]$shp.OnAction")


def test_43_re_reading_count_every_iteration_is_refused() -> None:
    _mutate("build", "test_3",
            "        for ($i = 1; $i -le $compCount; $i++) {",
            "        for ($i = 1; $i -le $vbcomps2.Count; $i++) {")


def test_44_wrapping_a_write_in_the_retry_is_refused() -> None:
    """REQUIRED: do not broadly wrap all Excel automation in retries."""
    _mutate("build", "test_3",
            "    $wb.Save()\n",
            "    $null = Invoke-ComRetryRead -Target $wb -Member 'Save' -Description 'save'\n")


def test_45_retrying_the_workbook_open_is_refused() -> None:
    """OPEN IS NOT IDEMPOTENT: a second attempt opens a second copy."""
    _mutate("build", "test_34",
            "        $wb2 = $workbooks2.Open($stageBPath)",
            "        $wb2 = (Invoke-ComRetryRead -Target $workbooks2 -Member 'Open' `\n"
            "                    -Description 'reopen').Open($stageBPath)")


def test_46_retrying_a_release_or_shutdown_call_is_refused() -> None:
    _mutate("build", "test_35",
            "        $compCount = [int](Invoke-ComRetryRead -Target $vbcomps2 -Member 'Count' `",
            "        $null = Invoke-ComRetryRead -Target $wb2 -Member 'Close' -Description 'x'\n"
            "        $compCount = [int](Invoke-ComRetryRead -Target $vbcomps2 -Member 'Count' `")


# ===========================================================================
# F. COVERAGE AND REPORTING
# ===========================================================================
def test_50_dropping_the_button_verification_is_refused() -> None:
    """REQUIRED: 11 button persistence checks may not be traded for stability."""
    _mutate("build", "test_40",
            "        foreach ($button in $manifest.buttons) {",
            "        foreach ($button in @()) {")


def test_51_dropping_the_module_verification_is_refused() -> None:
    _mutate("build", "test_40",
            "        foreach ($m in $manifest.vba.modules) {",
            "        foreach ($m in @()) {")


def test_52_dropping_the_codename_verification_is_refused() -> None:
    _mutate("build", "test_40",
            "        foreach ($sheet in $manifest.sheets) {",
            "        foreach ($sheet in @()) {")


def test_53_suppressing_events_to_make_verification_easier_is_refused() -> None:
    """THE VERIFICATION INSTANCE MUST OPEN THE WORKBOOK AS IT WILL FOR A PERSON."""
    _mutate("build", "test_42",
            "        $excel2.AskToUpdateLinks = $false",
            "        $excel2.AskToUpdateLinks = $false\n        $excel2.EnableEvents = $false")


def test_54_unprotecting_the_workbook_during_verification_is_refused() -> None:
    _mutate("build", "test_42",
            "        $worksheets2 = (Invoke-ComRetryRead -Target $wb2 -Member 'Worksheets' `",
            "        $wb2.Unprotect()\n"
            "        $worksheets2 = (Invoke-ComRetryRead -Target $wb2 -Member 'Worksheets' `")


def test_55_dropping_the_retry_report_is_refused() -> None:
    """REQUIRED: total retry duration finite AND reported."""
    _mutate("build", "test_50",
            "    $retryLines = @(Get-ComRetryLedger)",
            "    $retryLines = @()")


def test_56_hiding_the_wait_total_is_refused() -> None:
    _mutate("build", "test_50",
            "ms waited in total\" -f $retryLines.Count, (Get-ComRetryWaitTotal))",
            "read(s) were reissued\" -f $retryLines.Count)")


def test_57_reporting_the_retry_before_verification_concludes_is_refused() -> None:
    """A STEP THAT RAN FIRST COULD NOT KNOW WHETHER VERIFICATION FAILED."""
    _mutate("build", "test_52",
            "        $problems = @()\n",
            "        $problems = @()\n"
            "        Add-Step 'Transient COM rejections' 'PASS' 'nothing yet'\n")


def test_58_dropping_strict_mode_is_refused() -> None:
    _mutate("lifecycle", "test_55",
            "Set-StrictMode -Version 2.0",
            "Set-StrictMode -Off")


def test_59_deleting_the_helper_is_refused() -> None:
    """THE REWRITE-DELETED-A-FUNCTION CONTROL. It happened once already."""
    _mutate("lifecycle", "test_56",
            "function Get-ComRetryWaitTotal",
            "function Get-ComRetryWaitTotalRenamed")


def test_60_dropping_the_policy_declaration_is_refused() -> None:
    """A CAPABILITY IN THE CODE AND NOT IN THE POLICY IS ONE NOBODY AGREED TO."""
    _mutate("lifecycle", "test_54",
            "#   * A REFUSED call may be reissued; a FAILED one may not.",
            "#   * Retries happen where needed.")


# ===========================================================================
# G. THE DECLARATION THAT WIDENED THE BYTE FREEZE
# ===========================================================================
def test_70_declaring_a_frozen_scenario_harness_changeable_is_refused() -> None:
    """test_56 WIDENED TWICE BY DECLARATION. This proves the declaration cannot
    be widened to cover a harness the freeze exists to protect."""
    import test_phase10_benchmark_harness as bench

    saved = set(bench.CHANGED_BY_DECLARATION)
    try:
        bench.CHANGED_BY_DECLARATION.add("phase7_w1_smoke.ps1")
        failed = False
        try:
            bench.test_56b_the_declaration_and_the_freeze_do_not_overlap()
        except AssertionError:
            failed = True
        assert failed, "a frozen harness was admitted to the declaration"
    finally:
        bench.CHANGED_BY_DECLARATION.clear()
        bench.CHANGED_BY_DECLARATION.update(saved)


def test_71_a_new_unaccounted_harness_is_refused() -> None:
    """A SCENARIO HARNESS ADDED AND LEFT OFF BOTH LISTS WOULD BE UNPROTECTED."""
    import test_phase10_benchmark_harness as bench

    saved = tuple(bench.FROZEN_HARNESSES)
    try:
        bench.FROZEN_HARNESSES = saved[:-1]
        failed = False
        try:
            bench.test_56b_the_declaration_and_the_freeze_do_not_overlap()
        except AssertionError:
            failed = True
        assert failed, "a harness on neither list passed unnoticed"
    finally:
        bench.FROZEN_HARNESSES = saved


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-q"]))
