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
    # RE-ANCHORED, NOT LOOSENED. The ledger now has two readers and each reports
    # its own slice, so the statement this mutation damages moved with it.
    _mutate("build", "test_50",
            "    $retryLines = @(@(Get-ComRetryLedger) | Select-Object -Skip $verifyRetryBase)",
            "    $retryLines = @()")


def test_56_hiding_the_wait_total_is_refused() -> None:
    _mutate("build", "test_50",
            "ms waited in total\" -f $retryLines.Count, ((Get-ComRetryWaitTotal) - $verifyRetryWaitBase))",
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


# ===========================================================================
# G. THE BUILD-OPERATION INSTRUMENTATION AND THE FOUR OPTED-IN READS
# ===========================================================================
# THE NEW CONTROLS RUN A HARNESS AGAINST THE FILE ON DISK, so an in-memory
# mutation would never reach them: the damaged copy would sit in `_MEMO` while
# pwsh read the undamaged file and reported a clean sweep over the damage. These
# mutations therefore write the file, re-run the whole battery, and restore it in
# a `finally` - the same shape the reserved-row battery uses for the same reason.
_ON_DISK = {
    "lifecycle": conformance.LIFECYCLE_PS1,
    "build": conformance.BUILD_PS1,
    "gate": conformance.GATE_PS1,
    "evidence": conformance.EVIDENCE_MD,
}


def _mutate_on_disk(key: str, expected: str, before: str, after: str) -> None:
    """Damage one file ON DISK, re-run everything, and always put it back."""
    path = _ON_DISK[key]
    with path.open(encoding="utf-8", newline="") as handle:
        original = handle.read()
    crlf = "\r\n" in original
    damaged = original.replace(
        before.replace("\n", "\r\n") if crlf else before,
        after.replace("\n", "\r\n") if crlf else after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:70]!r} is no longer in {key}")
    try:
        with path.open("w", encoding="utf-8", newline="") as handle:
            handle.write(damaged)
        conformance._MEMO.clear()
        conformance._FLOW.clear()
        refused = _run_battery()
    finally:
        with path.open("w", encoding="utf-8", newline="") as handle:
            handle.write(original)
        conformance._MEMO.clear()
        conformance._FLOW.clear()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def test_80_retrying_every_com_exception_is_refused_when_executed() -> None:
    """THE NAMED MUTATION, CAUGHT BY BEHAVIOUR RATHER THAN BY TEXT. Retrying on the
    exception TYPE turns "the call never ran" into "something went wrong, try
    again" - and the harness then reissues an error Excel ACCEPTED."""
    _mutate_on_disk(
        "lifecycle", "test_6",
        "            $name = Get-ComRejectionName $_\n",
        "            $name = ''\n"
        "            if ($_.Exception -is [System.Runtime.InteropServices.COMException]) "
        "{ $name = 'any COM failure' }\n"
        "            if ($null -ne $_.Exception.InnerException -and "
        "$_.Exception.InnerException -is [System.Runtime.InteropServices.COMException]) "
        "{ $name = 'any COM failure' }\n")


def test_81_removing_the_attempt_bound_is_refused() -> None:
    """A LIMIT THAT IS NEVER CONSULTED IS NOT A LIMIT. With the attempt reason gone,
    only the budget can end the loop, and an always-refused read is reissued far
    past the three attempts it was allowed."""
    _mutate_on_disk(
        "lifecycle", "test_6",
        "            if ($attempts -ge $MaxAttempts) {\n"
        "                $reason = ('attempt limit of ' + [string]$MaxAttempts + ' reached')\n"
        "            } elseif",
        "            if ($false) {\n"
        "                $reason = ('attempt limit of ' + [string]$MaxAttempts + ' reached')\n"
        "            } elseif")


def test_82_removing_the_total_wait_bound_is_refused() -> None:
    """THE OTHER BOUND, AND IT IS TESTED SEPARATELY BECAUSE A SMALL ATTEMPT LIMIT
    WOULD OTHERWISE HIDE IT. With the budget check gone, a five-millisecond budget
    stops meaning anything and the read is reissued up to its attempt limit."""
    _mutate_on_disk(
        "lifecycle", "test_67",
        "            } elseif (($waitedMs + $delay) -gt $TotalBudgetMs) {",
        "            } elseif ($false) {")


def test_83_swallowing_an_exhausted_rejection_is_refused() -> None:
    """THE WORST FAILURE MODE AVAILABLE HERE. A retry that gives up and RETURNS
    hands the caller a $null it never read, and the bootstrap reports state it did
    not observe."""
    _mutate_on_disk(
        "lifecycle", "test_6",
        "                $script:comRetryWaitMs = $script:comRetryWaitMs + $waitedMs\n"
        "                throw\n",
        "                $script:comRetryWaitMs = $script:comRetryWaitMs + $waitedMs\n"
        "                $answered = $true\n"
        "                break\n")


def test_84_routing_saveas_through_the_read_retry_is_refused() -> None:
    """RE-ANCHORED, NOT LOOSENED. SaveAs is still not idempotent; what changed is
    that it now has a POSTCONDITION-GATED recovery of its own, so the mutation that
    matters is routing it through the blind READ retry instead - a reissue with no
    observation behind it, which is exactly the guess the refusal contract does not
    license."""
    _mutate_on_disk(
        "build", "test_",
        "            $Workbook.SaveAs($TargetPath, $TargetFormat)",
        "            $null = (Invoke-StageBBuildRead -Target $Workbook -Member 'SaveAs' `\n"
        "                         -Operation 'saveas.xlsm' -Description 'SaveAs').Value")


def test_85_routing_the_module_import_through_the_retry_is_refused() -> None:
    """A REISSUED Import CAN LEAVE modConstants1 BESIDE modConstants. The build's
    own comment says so, and Excel will do it without complaint."""
    _mutate_on_disk(
        "build", "test_7",
        "            $imported = $vbcomps.Import($file)",
        "            $imported = (Invoke-StageBBuildRead -Target $vbcomps -Member 'Import' `\n"
        "                             -Operation 'vbcomponents.import' -Description 'Import').Value")


def test_86_routing_the_button_creation_through_the_retry_is_refused() -> None:
    """A REISSUED AddShape LEAVES TWO BUTTONS. The manifest declares eleven, and a
    twelfth with the same OnAction would still read back correctly."""
    _mutate_on_disk(
        "build", "test_7",
        "            $shp = $shapes.AddShape(5, [double]$anchor.Left, [double]$anchor.Top, [double]$button.width, [double]$button.height)",
        "            $shp = (Invoke-StageBBuildRead -Target $shapes -Member 'AddShape' `\n"
        "                        -Operation 'button.add' -Description 'AddShape').Value")


def test_87_omitting_the_operation_from_the_failure_is_refused() -> None:
    """THE ENTIRE POINT OF THE BATCH. Without the label the transcript is back to
    naming a region, which is what four runs already produced."""
    _mutate_on_disk(
        "build", "test_62",
        "    Add-Step 'Stage-B build' 'FAIL' ('operation=' + $failedOp + '; ' + (Format-Err $_))",
        "    Add-Step 'Stage-B build' 'FAIL' (Format-Err $_)")


def test_88_dropping_a_label_before_a_named_operation_is_refused() -> None:
    """A LABEL THAT IS NEVER SET CANNOT APPEAR IN A DIAGNOSTIC, and the operation
    before it would be blamed for this one's refusal."""
    _mutate_on_disk(
        "build", "test_",
        "    Set-StageBBuildOp 'workbook.save'\n    $wb.Save()",
        "    $wb.Save()")


def test_89_opening_the_vocabulary_to_any_string_is_refused() -> None:
    """A CLOSED VOCABULARY IS WHAT MAKES A LABEL TRACEABLE. If any string is
    accepted, a misspelling reaches the transcript and maps to no call site."""
    _mutate_on_disk(
        "build", "test_61",
        "    if ($script:StageBBuildOps -notcontains $Operation) {",
        "    if ($false) {")


def test_90_setting_the_label_after_the_read_is_refused() -> None:
    """A READ REFUSED ON ITS FIRST ATTEMPT WOULD CARRY THE PREVIOUS STEP'S NAME - a
    diagnostic that is confidently wrong, which is worse than none.

    RE-ANCHORED. The wrapper this used to damage is gone: run 5 proved it returned
    no value on Windows. The property is unchanged and now lives at the read itself,
    so the mutation moves the label to AFTER the call it is supposed to name.
    """
    # RE-ANCHORED AGAIN, and to the readiness gate this time: the pre-save reads moved
    # into it when Windows proved the post-open gap. The property is unchanged.
    _mutate_on_disk(
        "build", "test_69",
        "        Set-StageBBuildStep 'open.ready.fileformat'\n"
        "        $formatValue = $null\n"
        "        try {\n"
        "            $formatRead = Invoke-ComRetryRead -Target $Workbook -Member 'FileFormat' `\n"
        "                              -Description 'the opened workbook FileFormat'",
        "        $formatValue = $null\n"
        "        try {\n"
        "            $formatRead = Invoke-ComRetryRead -Target $Workbook -Member 'FileFormat' `\n"
        "                              -Description 'the opened workbook FileFormat'\n"
        "            Set-StageBBuildStep 'open.ready.fileformat'")


def test_91_suppressing_the_clean_run_line_is_refused() -> None:
    """SILENCE AND SUCCESS LOOKED IDENTICAL IN ALL FOUR RUNS. "Nothing was refused"
    has to be printed, or a build that waited nine seconds reads like a quiet one."""
    _mutate_on_disk(
        "build", "test_79",
        "    Add-Step 'Transient COM rejections (build)' 'PASS' 'COMREJECT|build|none|attempts=0|waited=0'",
        "    $null = 'no line'")


def test_92_a_blanket_sleep_in_the_bootstrap_is_refused() -> None:
    """THE FIX IS NOT "WAIT A BIT AND HOPE". A sleep after the open would be a
    readiness gate under another name, and it is not authorised."""
    # RE-ANCHORED to the workbook open, which is where a readiness sleep would
    # actually be put. The old anchor was the SaveAs statement that has since moved
    # into its own function.
    _mutate_on_disk(
        "build", "test_",
        "    Set-StageBBuildOp 'open.workbook'\n"
        "    $workbooks = $excel.Workbooks",
        "    Start-Sleep -Milliseconds 1500\n"
        "    Set-StageBBuildOp 'open.workbook'\n"
        "    $workbooks = $excel.Workbooks")


def test_93_an_inter_pass_drain_in_the_gate_is_refused() -> None:
    """NOT AUTHORISED, AND IT WOULD HIDE THE CAUSE. Both instances shut down
    naturally; a pause between the passes treats a symptom nobody has explained."""
    _mutate_on_disk(
        "gate", "test_76",
        "    $bootstrapExit = $LASTEXITCODE",
        "    Start-Sleep -Seconds 20\n"
        "    $bootstrapExit = $LASTEXITCODE")


def test_94_dropping_the_bootstrap_exit_check_is_refused() -> None:
    """THE HALF-BUILT WORKBOOK. Without it, a build refused after SaveAs leaves an
    .xlsm with no modules and the pass reports a fixture result against it."""
    _mutate_on_disk(
        "gate", "test_78",
        "    if ($bootstrapExit -ne 0) {",
        "    if ($false) {")


def test_95_claiming_the_historical_failing_call_is_refused() -> None:
    """THE EVIDENCE RULE. Static work cannot name the call runs 3 and 4 were refused
    on. A record that says it was VBProject would licence a correction built on a
    guess."""
    _mutate_on_disk(
        "evidence", "test_85",
        "**`VBProject` acquisition is a CANDIDATE and nothing more.**",
        "**The rejected call was `VBProject` acquisition.** the failing call was identified.")


def test_96_softening_the_runs_3_and_4_verdict_is_refused() -> None:
    """A RECORD THAT LET THEM READ AS A SEMANTIC RESULT WOULD LICENCE A BULK
    BASELINE ON EVIDENCE THAT DOES NOT EXIST."""
    # THE ANCHOR IS THE RUNS-3/4 SENTENCE, not the phrase. Run 2's record carries
    # the same words, and a first-occurrence replacement damaged that section
    # instead - the mutation then survived because it was aimed at the wrong one.
    _mutate_on_disk(
        "evidence", "test_85",
        "No semantic comparison was executed. This record\n"
        "**must not be read as a fixture DIFFER**.",
        "The Bulk pass differed from Endpoints.")


def test_97_erasing_the_missing_identification_from_the_record_is_refused() -> None:
    """THE FINDING IS THAT THE LOG COULD NOT SAY. Dropping that sentence leaves a
    record in which four runs simply failed, and the reason for this batch
    disappears."""
    _mutate_on_disk(
        "evidence", "test_85",
        "### The old Stage-B log did not identify the exact rejected COM operation",
        "### The old Stage-B log")


# ===========================================================================
# H. THE SaveAs SETTLEMENT
# ===========================================================================
# WINDOWS NAMED saveas.xlsm, AND SaveAs WRITES A FILE. Every way of turning
# "observe, then decide" back into "decide, then hope" is fed in here, and the
# executed harness - which counts the actual SaveAs calls - is what catches it.
def test_98_blindly_reissuing_the_save_is_refused() -> None:
    """THE WHOLE SETTLEMENT IN ONE MUTATION. Retrying on the refusal contract alone
    reissues a save that may already have written the workbook."""
    _mutate_on_disk(
        "build", "test_9",
        "            $state = Get-StageBSaveAsPostcondition -Workbook $Workbook `\n"
        "                -SourcePath $SourcePath -TargetPath $TargetPath `\n"
        "                -TargetFormat $TargetFormat -SourceFormat $SourceFormat\n"
        "            Add-Note ('SAVEAS|postcondition|' + $state.State + '|' + $state.Detail)\n"
        "            if ($state.State -eq 'completed') {",
        "            $state = [pscustomobject]@{ State = 'not-executed'; FullName = ''; "
        "FileFormat = 0; BoundToTarget = $false; BoundToSource = $true; "
        "TargetExists = $false; SourceExists = $true; ReadError = ''; Detail = 'assumed' }\n"
        "            Add-Note ('SAVEAS|postcondition|' + $state.State + '|' + $state.Detail)\n"
        "            if ($state.State -eq 'completed') {")


def test_99_treating_the_target_file_alone_as_a_completed_save_is_refused() -> None:
    """A FILE ON DISK IS WHAT A HALF-WRITTEN SAVE LOOKS LIKE TOO. Accepting it
    would carry a workbook the build does not own into the module import."""
    _mutate_on_disk(
        "build", "test_9",
        "    } elseif ($boundToTarget -and ($format -eq $TargetFormat) -and $targetExists) {",
        "    } elseif ($targetExists) {")


def test_100_treating_a_changed_path_alone_as_a_completed_save_is_refused() -> None:
    """A REBOUND WORKBOOK WITH NO FILE IS NOT A SAVE."""
    _mutate_on_disk(
        "build", "test_9",
        "    } elseif ($boundToTarget -and ($format -eq $TargetFormat) -and $targetExists) {",
        "    } elseif ($boundToTarget) {")


def test_101_treating_a_changed_format_alone_as_a_completed_save_is_refused() -> None:
    """THE FORMAT MOVING IN MEMORY IS NOT THE FILE BEING WRITTEN."""
    _mutate_on_disk(
        "build", "test_9",
        "    } elseif ($boundToTarget -and ($format -eq $TargetFormat) -and $targetExists) {",
        "    } elseif ($format -eq $TargetFormat) {")


def test_102_retrying_an_ambiguous_save_is_refused() -> None:
    """THE STATE THAT MUST NEVER REISSUE. If the save can be proved neither to have
    happened nor to have been skipped, a second attempt is a coin toss with the
    workbook."""
    _mutate_on_disk(
        "build", "test_9",
        "            if ($state.State -ne 'not-executed') {\n"
        "                throw ('SAVEAS AMBIGUOUS after ",
        "            if ($false) {\n"
        "                throw ('SAVEAS AMBIGUOUS after ")


def test_103_retrying_a_non_rpc_com_error_is_refused() -> None:
    """AN ERROR EXCEL ACCEPTED DESCRIBES SOMETHING THAT HAPPENED. Reissuing the
    save over it is exactly the guess the refusal contract does not license."""
    _mutate_on_disk(
        "build", "test_",
        "            if ([string]::IsNullOrWhiteSpace($refused)) {\n"
        "                # EXCEL ACCEPTED THIS ONE AND IT FAILED. That describes something\n"
        "                # which actually happened, and nothing here may reissue it.\n"
        "                Add-Note ('SAVEAS|attempt=' + [string]$attempt + '|error|' + $hres)\n"
        "                throw\n"
        "            }",
        "            if ([string]::IsNullOrWhiteSpace($refused)) {\n"
        "                Add-Note ('SAVEAS|attempt=' + [string]$attempt + '|error|' + $hres)\n"
        "                $refused = 'treated as refused'\n"
        "            }")


def test_104_reissuing_a_save_that_completed_is_refused() -> None:
    """IT WOULD OVERWRITE THE WORKBOOK THIS BUILD ALREADY OWNS."""
    _mutate_on_disk(
        "build", "test_9",
        "            if ($state.State -eq 'completed') {",
        "            if ($false) {")


def test_105_dropping_the_fullname_verification_is_refused() -> None:
    """WITHOUT IT, A SAVE THAT WROTE THE FILE BUT LEFT THE WORKBOOK ON THE SOURCE
    READS AS COMPLETE - and every later step writes into the wrong workbook."""
    _mutate_on_disk(
        "build", "test_",
        "    $boundToTarget = ($seen -ne '') -and ($seen -eq (Get-StageBComparablePath $TargetPath))",
        "    $boundToTarget = $true")


def test_106_dropping_the_fileformat_verification_is_refused() -> None:
    """A .xlsx SAVED UNDER AN .xlsm NAME HAS NO VBA PROJECT, and the import would
    fail three steps later with nothing pointing back to here."""
    _mutate_on_disk(
        "build", "test_",
        "    } elseif ($boundToTarget -and ($format -eq $TargetFormat) -and $targetExists) {",
        "    } elseif ($boundToTarget -and $targetExists) {")


def test_107_dropping_the_target_existence_verification_is_refused() -> None:
    """THE WORKBOOK CAN CLAIM A PATH IT NEVER WROTE."""
    _mutate_on_disk(
        "build", "test_",
        "    $targetExists  = [bool](Test-Path -LiteralPath $TargetPath)",
        "    $targetExists  = $true")


def test_108_removing_the_save_attempt_bound_is_refused() -> None:
    """A BOUND THAT IS NEVER CONSULTED IS NOT A BOUND, and an unbounded SaveAs loop
    on a Windows host holds an Excel process open and produces no transcript."""
    _mutate_on_disk(
        "build", "test_97",
        "            if ($attempt -ge $MaxAttempts) { throw }",
        "            if ($false) { throw }")


def test_109_removing_the_save_wait_bound_is_refused() -> None:
    """THE OTHER BOUND, TESTED SEPARATELY because a small attempt limit would
    otherwise hide it."""
    _mutate_on_disk(
        "build", "test_97",
        "            if (($waitedMs + $delay) -gt $TotalBudgetMs) { throw }",
        "            if ($false) { throw }")


def test_110_swallowing_an_exhausted_save_failure_is_refused() -> None:
    """THE WORST OUTCOME AVAILABLE. A save that gave up and RETURNED would let the
    module import run against a workbook that was never written."""
    _mutate_on_disk(
        "build", "test_",
        "    throw ('Invoke-StageBSaveAs: the save was never completed after ' +\n"
        "           [string]$attempt + ' attempt(s).')",
        "    return (New-StageBSaveAsResult -State ([pscustomobject]@{ State = 'completed'; "
        "FullName = $TargetPath; FileFormat = $TargetFormat; Detail = 'assumed' }) "
        "-Attempts $attempt -WaitedMs $waitedMs)")


def test_111_routing_the_module_import_through_the_save_retry_is_refused() -> None:
    """THE SETTLEMENT IS FOR ONE CALL. Import has no postcondition of this kind and
    a reissued one leaves modConstants1 beside modConstants."""
    _mutate_on_disk(
        "build", "test_101",
        "            $imported = $vbcomps.Import($file)",
        "            $imported = Invoke-StageBSaveAs -Workbook $vbcomps -SourcePath $file `\n"
        "                            -TargetPath $file -TargetFormat 52 -SourceFormat 51")


def test_112_accepting_a_normal_return_without_verifying_it_is_refused() -> None:
    """A COM METHOD THAT RETURNED IS NOT A SAVE THAT HAPPENED. This is the quiet
    path, and it is the one a reader is most likely to trust."""
    _mutate_on_disk(
        "build", "test_89",
        "        if ($state.State -ne 'completed') {\n"
        "            throw ('SaveAs returned without error but its postconditions do not prove ' +\n"
        "                   'the save: ' + $state.Detail)\n"
        "        }",
        "        if ($false) {\n"
        "            throw ('SaveAs returned without error but its postconditions do not prove ' +\n"
        "                   'the save: ' + $state.Detail)\n"
        "        }")


def test_113_a_blanket_sleep_outside_the_save_retry_is_refused() -> None:
    """THE SLEEP IS GATED ON AN OBSERVATION. One anywhere else is the readiness gate
    this batch is not authorised to add."""
    _mutate_on_disk(
        "build", "test_",
        "    Set-StageBBuildOp 'vbcomponents.import'",
        "    Start-Sleep -Milliseconds 2000\n"
        "    Set-StageBBuildOp 'vbcomponents.import'")


def test_114_ungating_the_backoff_from_the_observation_is_refused() -> None:
    """SLEEPING BEFORE THE STATE IS KNOWN turns the settlement back into a wait."""
    _mutate_on_disk(
        "build", "test_14",
        "            Add-Note ('SAVEAS|attempt=' + [string]$attempt + '|rejected|' + $hres)\n"
        "            # INSPECT BEFORE ANY SECOND CALL. This is the whole settlement.",
        "            Add-Note ('SAVEAS|attempt=' + [string]$attempt + '|rejected|' + $hres)\n"
        "            Start-Sleep -Milliseconds 50\n"
        "            # INSPECT BEFORE ANY SECOND CALL. This is the whole settlement.")


def test_115_cleaning_up_an_ambiguous_save_is_refused() -> None:
    """NO SILENT CLEANUP THAT DESTROYS EVIDENCE. The half-written file is the only
    record of what happened."""
    _mutate_on_disk(
        "build", "test_93",
        "                throw ('SAVEAS AMBIGUOUS after ' + [string]$attempt + ' attempt(s): the ' +",
        "                Remove-Item -LiteralPath $TargetPath -Force -ErrorAction SilentlyContinue\n"
        "                throw ('SAVEAS AMBIGUOUS after ' + [string]$attempt + ' attempt(s): the ' +")


def test_116_hard_coding_the_source_format_is_refused() -> None:
    """'NOTHING MOVED' IS ONLY PROVABLE AGAINST WHAT THE WORKBOOK WAS. A literal 51
    here is this script restating a contract it is supposed to read."""
    _mutate_on_disk(
        "build", "test_99",
        "    } elseif ($boundToSource -and ($format -eq $SourceFormat) -and (-not $targetExists)) {",
        "    } elseif ($boundToSource -and ($format -eq 51) -and (-not $targetExists)) {")


def test_117_removing_the_pre_save_target_deletion_is_refused() -> None:
    """IT IS WHAT MAKES THE OBSERVATION CONCLUSIVE. With a stale target left in
    place, 'the file is there' stops meaning 'this call put it there'."""
    _mutate_on_disk(
        "build", "test_100",
        "    if (Test-Path -LiteralPath $stageBPath) { Remove-Item -LiteralPath $stageBPath -Force }",
        "    $null = 'the stale target is left alone'")


def test_118_speculating_about_the_root_cause_in_the_record_is_refused() -> None:
    """WE KNOW WHERE, NOT WHY. A record that names a cause nobody proved would send
    the next batch after the wrong thing."""
    _mutate_on_disk(
        "evidence", "test_104",
        "**We know WHERE the rejection occurs. We do not know WHY Excel rejects the\n"
        "second-session `SaveAs`.**",
        "**The cause is a lifecycle overlap between the two Excel sessions.**")


# ===========================================================================
# I. THE READ THAT ANSWERED WITH NOTHING
# ===========================================================================
# RUN 5 FAILED IN THE PRE-SAVE OBSERVATION AND CALLED ITSELF saveas.xlsm. Every way
# of losing a value, polluting it, accepting its absence, or mislabelling where it
# happened is fed in here.
def test_119_suppressing_the_returned_fileformat_is_refused() -> None:
    """RUN 5's DEFECT, PLANTED. A read whose value never arrives must fail at the
    operation it happened at - not three statements later on a garbage baseline."""
    # RE-ANCHORED: the baseline now takes the value the readiness gate observed.
    _mutate_on_disk(
        "build", "test_",
        "    $sourceFormat = Get-StageBScalarInt -Value $ready.FileFormat `\n"
        "                        -What 'the Stage-A workbook FileFormat observed at open'",
        "    $sourceFormat = Get-StageBScalarInt -Value $null `\n"
        "                        -What 'the Stage-A workbook FileFormat observed at open'")


def test_120_emitting_telemetry_before_the_value_is_refused() -> None:
    """ASSIGN, NEVER EMIT. A recorder that writes to the output stream puts its line
    in front of the value, and the caller reads the line as the answer."""
    _mutate_on_disk(
        "build", "test_",
        "    $null = $script:StageBBuildRejections.Add(\n",
        "    Write-Output ('COMREJECT|build|' + $Operation)\n"
        "    $null = $script:StageBBuildRejections.Add(\n")


def test_121_emitting_telemetry_after_the_value_is_refused() -> None:
    """THE SAME DEFECT FROM THE OTHER SIDE: a trailing emission makes the return an
    ARRAY, and `.Value` on an array is not the value."""
    _mutate_on_disk(
        "build", "test_",
        "function Add-StageBReadRejection {\n"
        "    param([string]$Operation, $Record)\n"
        "    if ($null -eq $Record) { return }",
        "function Add-StageBReadRejection {\n"
        "    param([string]$Operation, $Record)\n"
        "    $Operation\n"
        "    if ($null -eq $Record) { return }")


def test_122_treating_a_null_answer_as_a_value_is_refused() -> None:
    """A READ THAT NEITHER RAISED NOR ANSWERED IS NOT AN ANSWER. [int]$null is 0,
    and 0 as a source FileFormat makes every NOT-EXECUTED verdict wrong."""
    _mutate_on_disk(
        "build", "test_82",
        "    if ($null -eq $Value) {\n"
        "        throw ($What + ' answered with nothing. The read was not refused and it was ' +\n"
        "               'not answered.')\n"
        "    }\n"
        "    if ($Value -is [System.Array]) {\n"
        "        throw ($What + ' answered with ' + [string]@($Value).Count + ' values where one ' +\n"
        "               'scalar was required.')\n"
        "    }",
        "    if ($null -eq $Value) { return 0 }")


def test_123_using_truthiness_instead_of_a_null_test_is_refused() -> None:
    """`if (-not $value)` REFUSES A LEGITIMATE ZERO. FileFormat is numeric and 0 is a
    real value elsewhere in the object model, so the test is for absence and type,
    never for truth."""
    _mutate_on_disk(
        "build", "test_82",
        "function Get-StageBScalarInt {\n"
        "    param($Value, [string]$What)\n"
        "    if ($null -eq $Value) {",
        "function Get-StageBScalarInt {\n"
        "    param($Value, [string]$What)\n"
        "    if (-not $Value) { throw ($What + ' answered nothing useful.') }\n"
        "    if ($null -eq $Value) {")


def test_124_accepting_a_non_scalar_fileformat_is_refused() -> None:
    """TWO VALUES IS NOT ONE. An array answer compared against $SourceFormat would
    be neither equal nor unequal in a way anyone can reason about."""
    _mutate_on_disk(
        "build", "test_82",
        "    if ($Value -is [System.Array]) {\n"
        "        throw ($What + ' answered with ' + [string]@($Value).Count + ' values where one ' +\n"
        "               'scalar was required.')\n"
        "    }\n"
        "    $text = ([string]$Value).Trim()",
        "    $text = ([string]$Value).Trim()")


def test_125_accepting_an_empty_fullname_is_refused() -> None:
    """AN EMPTY PATH COMPARES EQUAL TO NOTHING, so boundToSource and boundToTarget
    would both be false and every state would read as AMBIGUOUS."""
    _mutate_on_disk(
        "build", "test_82",
        "    if ([string]::IsNullOrWhiteSpace($text)) {\n"
        "        throw ($What + ' answered an empty string.')\n"
        "    }",
        "    if ($false) {\n"
        "        throw ($What + ' answered an empty string.')\n"
        "    }")


def test_126_saving_without_a_consistent_baseline_is_refused() -> None:
    """NOT EXECUTED CANNOT BE RECOGNISED WITHOUT A BASELINE. If the workbook is not
    the Stage-A file, 'bound to source' means nothing."""
    _mutate_on_disk(
        "build", "test_99",
        "    if ($preSeen -ne (Get-StageBComparablePath $stageAPath)) {",
        "    if ($false) {")


def test_127_saving_over_an_existing_target_is_refused() -> None:
    """IF THE TARGET IS ALREADY THERE, 'the file exists' afterwards proves nothing -
    which is the whole basis of the COMPLETED verdict."""
    _mutate_on_disk(
        "build", "test_99",
        "    if (Test-Path -LiteralPath $stageBPath) {\n"
        "        throw ('SAVEAS BASELINE: ' + $stageBPath + ' is present before the save, so ' +",
        "    if ($true -and $false) {\n"
        "        throw ('SAVEAS BASELINE: ' + $stageBPath + ' is present before the save, so ' +")


def test_128_labelling_a_pre_save_read_as_the_save_is_refused() -> None:
    """RUN 5's REPORTING DEFECT. A failed observation reported as saveas.call reads
    as though the save had been attempted, and the next batch would chase the wrong
    thing."""
    # RE-ANCHORED to the readiness gate. Labelling its FileFormat observation
    # saveas.call reads as though the save had been attempted, which is run 5's
    # reporting defect arriving one function earlier.
    _mutate_on_disk(
        "build", "test_10",
        "        Set-StageBBuildStep 'open.ready.fileformat'",
        "        Set-StageBBuildStep 'saveas.call'")


def test_129_reporting_the_region_instead_of_the_step_is_refused() -> None:
    """THE COARSE LABEL IS WHAT RUN 5 PRINTED."""
    _mutate_on_disk(
        "build", "test_62",
        "    $failedOp = Get-StageBBuildLabel",
        "    $failedOp = Get-StageBBuildOp")


def test_130_letting_a_stale_sub_operation_survive_is_refused() -> None:
    """A SUB-OPERATION THAT ALREADY FINISHED WOULD NAME THE NEXT FAILURE."""
    _mutate_on_disk(
        "build", "test_107",
        "    $script:StageBBuildOp = $Operation\n"
        "    # A NEW TOP-LEVEL OPERATION CLEARS THE SUB-OPERATION. A stale one would name a\n"
        "    # step that finished for a failure somewhere else entirely.\n"
        "    $script:StageBBuildSubOp = ''",
        "    $script:StageBBuildOp = $Operation")


def test_131_opening_the_sub_operation_vocabulary_is_refused() -> None:
    """A CLOSED VOCABULARY IS WHAT MAKES A LABEL TRACEABLE."""
    _mutate_on_disk(
        "build", "test_107",
        "    if ($script:StageBBuildSubOps -notcontains $Step) {",
        "    if ($false) {")


def test_132_bringing_the_read_wrapper_back_is_refused() -> None:
    """IT PRODUCED NO ANSWER ON WINDOWS. Reinstating an intermediate reader puts the
    unproven hop back between the COM object and the member access."""
    # RE-ANCHORED to the readiness gate, where the first post-open read now lives.
    _mutate_on_disk(
        "build", "test_",
        "            $nameRead = Invoke-ComRetryRead -Target $Workbook -Member 'FullName' `\n"
        "                            -Description 'the opened workbook FullName'",
        "            function Invoke-StageBBuildRead {\n"
        "                param($Target, [string]$Member, [string]$Description)\n"
        "                return Invoke-ComRetryRead -Target $Target -Member $Member -Description $Description\n"
        "            }\n"
        "            $nameRead = Invoke-StageBBuildRead -Target $Workbook -Member 'FullName' `\n"
        "                            -Description 'the opened workbook FullName'")


def test_133_taking_the_value_inside_the_read_expression_is_refused() -> None:
    """THE PROVEN SHAPE CAPTURES THE RECORD FIRST. Collapsing it back into one
    expression is the form that produced no answer, and it also loses the
    out-of-band telemetry."""
    _mutate_on_disk(
        "build", "test_105",
        "    $wsRead = Invoke-ComRetryRead -Target $wb -Member 'Worksheets' `\n"
        "                  -Description 'the Stage-B workbook Worksheets collection'\n"
        "    Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $wsRead\n"
        "    if ($null -eq $wsRead.Value) {\n"
        "        throw 'the Stage-B workbook Worksheets collection answered with nothing.'\n"
        "    }\n"
        "    $worksheets = $wsRead.Value",
        "    $worksheets = (Invoke-ComRetryRead -Target $wb -Member 'Worksheets' `\n"
        "                       -Description 'the Stage-B workbook Worksheets collection').Value")


def test_134_weakening_the_three_state_settlement_is_refused() -> None:
    """SEMANTIC FREEZE. This batch fixed a read; the classification it feeds must be
    exactly the accepted one."""
    _mutate_on_disk(
        "build", "test_111",
        "    } elseif ($boundToSource -and ($format -eq $SourceFormat) -and (-not $targetExists)) {",
        "    } elseif ($boundToSource) {")


def test_135_recording_run_5_as_a_saveas_rejection_is_refused() -> None:
    """IT WAS NOT ONE. SaveAs was never called, and a record that says otherwise
    would send the next batch after a rejection that did not happen."""
    _mutate_on_disk(
        "evidence", "test_109",
        "**`SaveAs` itself was NOT executed.**",
        "**`SaveAs` was refused again.**")


def test_136_dropping_the_run_5_verdict_is_refused() -> None:
    """A RECORD THAT LET RUN 5 READ AS A RESULT WOULD LICENCE A BULK BASELINE ON
    EVIDENCE THAT DOES NOT EXIST."""
    _mutate_on_disk(
        "evidence", "test_109",
        "and it is **not** another SaveAs rejection. It is\n**INVALID / NOT EVALUATED**",
        "and it is a partial result. It is\n**a fixture difference**")


# ===========================================================================
# J. THE POST-OPEN READINESS GATE
# ===========================================================================
# THE GATE IS THE EASIEST THING HERE TO TURN INTO A SLEEP. Every way of making it
# always succeed, accept an absence, poll over a real answer, or wait when it should
# not is fed in, and the executed harness - which counts reads, attempts and
# milliseconds - is what catches it.
def test_137_a_gate_that_always_succeeds_is_refused() -> None:
    """THE WHOLE POINT IN ONE MUTATION. A gate that reports READY without observing
    anything puts the run back where c66e752 was."""
    _mutate_on_disk(
        "build", "test_11",
        "        if (($nameState -eq 'ok') -and ($formatState -eq 'ok')) { $ready = $true; break }",
        "        $ready = $true; break")


def test_138_accepting_a_null_fullname_is_refused() -> None:
    """A WORKBOOK THAT WILL NOT SAY WHICH FILE IT IS cannot establish the baseline the
    NOT-EXECUTED verdict depends on."""
    _mutate_on_disk(
        "build", "test_11",
        "        if ($null -eq $nameValue) {\n"
        "            $nameState = 'no-answer'",
        "        if ($null -eq $nameValue) {\n"
        "            $nameState = 'ok'")


def test_139_accepting_a_null_fileformat_is_refused() -> None:
    """[int]$null IS 0, AND 0 AS THE SOURCE FORMAT MAKES EVERY NOT-EXECUTED VERDICT
    WRONG. That is run 5's defect arriving through the gate instead."""
    _mutate_on_disk(
        "build", "test_11",
        "        if ($null -eq $formatValue) {\n"
        "            $formatState = 'no-answer'",
        "        if ($null -eq $formatValue) {\n"
        "            $formatState = 'ok'")


def test_140_polling_over_a_real_wrong_path_is_refused() -> None:
    """WAITING CANNOT CHANGE WHICH WORKBOOK THIS IS. Treating an identity failure as a
    delay spends the whole budget on a question already answered, and then reports
    'not ready' for a workbook that was never the right one."""
    _mutate_on_disk(
        "build", "test_118",
        "                throw ('READY: the opened workbook is bound to ' + [string]$nameValue +\n"
        "                       ' and not to ' + $ExpectedPath + '. Waiting cannot change which ' +\n"
        "                       'workbook this is.')",
        "                $nameState = 'no-answer'")


def test_141_polling_over_an_unusable_fileformat_is_refused() -> None:
    """A REAL BUT UNUSABLE ANSWER IS AN ERROR, not a readiness delay."""
    _mutate_on_disk(
        "build", "test_118",
        "            throw ('READY: the opened workbook FileFormat answered ' +\n"
        "                   ([string]$formatValue).Trim() + ', which is not an integer.')",
        "            $formatState = 'no-answer'")


def test_142_hard_coding_the_source_format_is_refused() -> None:
    """OBSERVED, NEVER ASSUMED. A literal 51 here is this script restating a contract
    it is supposed to read - and it would make a workbook in any other format look
    ready when it is not the workbook the build expects."""
    _mutate_on_disk(
        "build", "test_114",
        "            $format      = [int](([string]$formatValue).Trim())",
        "            $format      = 51")


def test_143_treating_a_real_error_as_not_ready_is_refused() -> None:
    """AN HRESULT EXCEL ACCEPTED DESCRIBES SOMETHING THAT HAPPENED. Polling over it
    hides a genuine failure behind fifteen seconds of waiting."""
    _mutate_on_disk(
        "build", "test_117",
        "            if ([string]::IsNullOrWhiteSpace((Get-ComRejectionName $_))) { throw }\n"
        "            $nameValue = $null",
        "            $nameValue = $null")


def test_144_removing_the_readiness_attempt_bound_is_refused() -> None:
    """AN UNBOUNDED POLL ON A WINDOWS HOST HOLDS AN EXCEL PROCESS OPEN and produces no
    transcript at all."""
    _mutate_on_disk(
        "build", "test_119",
        "        if ($attempt -ge $MaxAttempts) { break }\n"
        "        if (($waitedMs + $delay) -gt $TotalBudgetMs) { break }",
        "        if (($waitedMs + $delay) -gt $TotalBudgetMs) { break }")


def test_145_removing_the_readiness_wait_bound_is_refused() -> None:
    """THE OTHER BOUND, TESTED SEPARATELY because a small attempt limit would
    otherwise hide it."""
    _mutate_on_disk(
        "build", "test_119",
        "        if (($waitedMs + $delay) -gt $TotalBudgetMs) { break }\n"
        "        Start-Sleep -Milliseconds $delay",
        "        Start-Sleep -Milliseconds $delay")


def test_146_an_unconditional_post_open_sleep_is_refused() -> None:
    """THE FIX IS NOT 'WAIT A BIT AND HOPE'. A delay after every Open is the blanket
    sleep this project has refused from the first readiness run, and it would cost
    every clean session the same seconds."""
    _mutate_on_disk(
        "build", "test_",
        "    $ready = Wait-StageBWorkbookReady -Workbook $wb -ExpectedPath $stageAPath",
        "    Start-Sleep -Milliseconds 3000\n"
        "    $ready = Wait-StageBWorkbookReady -Workbook $wb -ExpectedPath $stageAPath")


def test_147_sleeping_after_successful_readiness_is_refused() -> None:
    """A WORKBOOK THAT ANSWERED IS NEVER WAITED ON. Moving the break after the backoff
    charges every clean run 250 ms for nothing."""
    _mutate_on_disk(
        "build", "test_115",
        "        if (($nameState -eq 'ok') -and ($formatState -eq 'ok')) { $ready = $true; break }",
        "        if (($nameState -eq 'ok') -and ($formatState -eq 'ok')) {\n"
        "            $ready = $true\n"
        "            Start-Sleep -Milliseconds $delay\n"
        "            $waitedMs = $waitedMs + $delay\n"
        "            break\n"
        "        }")


def test_148_an_inter_pass_sleep_in_the_gate_is_refused() -> None:
    """STILL FORBIDDEN. The gap is handled where it manifests, not by spacing the
    sessions out - and a pause between passes would hide the condition the gate is
    meant to observe."""
    _mutate_on_disk(
        "gate", "test_121",
        "    $bootstrapExit = $LASTEXITCODE",
        "    Start-Sleep -Seconds 15\n"
        "    $bootstrapExit = $LASTEXITCODE")


def test_149_running_the_gate_after_the_baseline_is_refused() -> None:
    """IT MUST GATE EVERYTHING. A readiness check that runs after the baseline has
    already been taken is a report, not a gate."""
    _mutate_on_disk(
        "build", "test_113",
        "    $ready = Wait-StageBWorkbookReady -Workbook $wb -ExpectedPath $stageAPath\n",
        "")


def test_150_letting_the_gate_return_a_verdict_instead_of_aborting_is_refused() -> None:
    """EXHAUSTION MUST STOP STAGE-B BEFORE SaveAs. A gate that returns anyway hands
    the baseline a format nobody observed."""
    _mutate_on_disk(
        "build", "test_113",
        "    if (-not $ready) {\n"
        "        Add-Note ('READY|open|exhausted|attempts=' + [string]$attempt + '|waited=' +\n"
        "                  [string]$waitedMs + '|fullname=' + $nameState + '|fileformat=' + $formatState)",
        "    if ($false) {\n"
        "        Add-Note ('READY|open|exhausted|attempts=' + [string]$attempt + '|waited=' +\n"
        "                  [string]$waitedMs + '|fullname=' + $nameState + '|fileformat=' + $formatState)")


def test_151_rereading_the_format_for_the_baseline_is_refused() -> None:
    """A SECOND READ IS A SECOND CHANCE TO ANSWER DIFFERENTLY than the value the
    NOT-EXECUTED verdict is measured against."""
    _mutate_on_disk(
        "build", "test_120",
        "    $sourceFormat = Get-StageBScalarInt -Value $ready.FileFormat `\n"
        "                        -What 'the Stage-A workbook FileFormat observed at open'",
        "    $reFormat = Invoke-ComRetryRead -Target $wb -Member 'FileFormat' `\n"
        "                    -Description 'the Stage-A workbook FileFormat again'\n"
        "    $sourceFormat = Get-StageBScalarInt -Value $reFormat.Value `\n"
        "                        -What 'the Stage-A workbook FileFormat observed at open'")


def test_152_widening_the_gate_into_a_health_probe_is_refused() -> None:
    """NOT A BROAD EXCEL PROBE. Every extra member is another thing that can answer
    with nothing, and another reason a clean session waits."""
    _mutate_on_disk(
        "build", "test_114",
        "        Set-StageBBuildStep 'open.ready.fileformat'\n"
        "        $formatValue = $null",
        "        Set-StageBBuildStep 'open.ready.fileformat'\n"
        "        $null = Invoke-ComRetryRead -Target $Workbook -Member 'Saved' `\n"
        "                    -Description 'the opened workbook Saved flag'\n"
        "        $formatValue = $null")


def test_153_claiming_a_cause_for_the_readiness_gap_is_refused() -> None:
    """WE KNOW WHERE, NOT WHY. A record that names a cause nobody proved would send
    the next batch after the wrong thing."""
    _mutate_on_disk(
        "evidence", "test_123",
        "**That says WHERE the gap is observed. It does not say WHY.**",
        "**The cause is a message-filter race in the second session.**")


def test_154_dropping_the_cross_run_comparison_is_refused() -> None:
    """THE TWO RUNS TOGETHER ARE WHAT LICENCE THE GATE. Either half alone reads as a
    member-specific or wrapper-specific defect."""
    _mutate_on_disk(
        "evidence", "test_123",
        "So the evidence no longer supports a member-specific or wrapper-specific diagnosis.",
        "So the FullName read is unreliable.")


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-q"]))
