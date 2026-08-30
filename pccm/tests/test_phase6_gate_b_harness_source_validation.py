#!/usr/bin/env python3
"""PCCM Phase 6 Step-13 MUTATION CONTROLS for the Gate-B harness source battery.

A conformance test that cannot fail proves nothing. Every control damages one of
the authorities Step 13 reads - `build/phase6_gate_b_inspection.json`,
`build/phase6_gate_b_cases.json` or `bootstrap/windows/phase6_gate_b_scenarios.ps1`
- reruns the WHOLE Step-13 static battery against the damaged copy, and requires
a NAMED detector among the refusers.

This matters more here than anywhere else in the project. The harness cannot be
executed on Linux, so these tests are the ONLY thing that can fail before a
Windows run; a battery that would pass over a moved `F21` or a hard-coded
address would be giving false assurance about a file nobody can run yet.

Nothing here writes to the repository: damaged copies live in a temporary
directory that is deleted on the way out, including on the exception path.

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
from contextlib import contextmanager
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase6_gate_b_harness_source as conformance  # noqa: E402

_INSPECTION = conformance.INSPECTION_PATH.read_text(encoding="utf-8")
_CASES = conformance.GATE_B_CASES_PATH.read_text(encoding="utf-8")
_PHASE6 = conformance.PHASE6.read_text(encoding="utf-8")
_HARNESS = conformance.HARNESS.read_text(encoding="utf-8")


def _conformance_tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 40, names
    return names


def _run_battery() -> list[str]:
    refused = []
    for name in _conformance_tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


@contextmanager
def _installed(inspection: str | None = None, cases: str | None = None,
               phase6: str | None = None, harness: str | None = None):
    saved = (conformance.INSPECTION_PATH, conformance.GATE_B_CASES_PATH,
             conformance.PHASE6, dict(conformance._CACHE), conformance.HARNESS)
    with tempfile.TemporaryDirectory(prefix="pccm-step13-mutation-") as name:
        temp = Path(name)
        conformance._CACHE.clear()
        try:
            for damaged, original, attribute in (
                (inspection, _INSPECTION, "INSPECTION_PATH"),
                (cases, _CASES, "GATE_B_CASES_PATH"),
                (phase6, _PHASE6, "PHASE6"),
                (harness, _HARNESS, "HARNESS"),
            ):
                if damaged is None:
                    continue
                assert damaged != original, f"the mutation changed nothing in {attribute}"
                target = temp / getattr(conformance, attribute).name
                target.write_text(damaged, encoding="utf-8")
                setattr(conformance, attribute, target)
            yield
        finally:
            conformance.INSPECTION_PATH = saved[0]
            conformance.GATE_B_CASES_PATH = saved[1]
            conformance.PHASE6 = saved[2]
            conformance.HARNESS = saved[4]
            conformance._CACHE.clear()
            conformance._CACHE.update(saved[3])


def _control(expected: str, **damage) -> None:
    with _installed(**damage):
        refused = _run_battery()
    assert refused, "the mutation survived the whole Step-13 battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _json_mutation(text: str, edit) -> str:
    document = json.loads(text)
    edit(document)
    return json.dumps(document, indent=2, sort_keys=False) + "\n"


def _swap(text: str, old: str, new: str, count: int = 1) -> str:
    assert text.count(old) == count, (old[:80], text.count(old))
    return text.replace(old, new)


def test_00_the_accepted_artefacts_pass_every_detector() -> None:
    with _installed():
        refused = _run_battery()
    assert refused == [], refused


# ===========================================================================
# A. The inspection projection - every address the harness will trust
# ===========================================================================
def test_01_the_pending_sidecar_moves_one_cell() -> None:
    """F21 is the durable recovery authority. A harness aimed one row off would
    write a marker production never reads and report a recovery that never
    happened."""
    def edit(document):
        document["sim_data"]["pending_auto_nonce"]["cell"] = "F22"
        document["sim_data"]["pending_auto_nonce"]["row"] = 22
    _control("test_18", inspection=_json_mutation(_INSPECTION, edit))


def test_02_the_pending_sidecar_moves_one_column() -> None:
    def edit(document):
        document["sim_data"]["pending_auto_nonce"]["cell"] = "G21"
        document["sim_data"]["pending_auto_nonce"]["column"] = "G"
    _control("test_18", inspection=_json_mutation(_INSPECTION, edit))


def test_03_the_final_commit_range_moves() -> None:
    """The one write the active bank moves inside. A wrong range would compare
    the restoration of a block that was never committed."""
    def edit(document):
        document["publication"]["final_commit_range"] = "D22:D29"
    _control("test_18", inspection=_json_mutation(_INSPECTION, edit))


def test_04_the_two_bank_columns_are_swapped() -> None:
    """Every per-bank read would then land in the other bank, and a scenario
    proving "the inactive bank was not published" would be reading the one that
    was."""
    def edit(document):
        identity = document["sim_data"]["run_identity"]["bank_value_columns"]
        identity["A"], identity["B"] = identity["B"], identity["A"]
    _control("test_18", inspection=_json_mutation(_INSPECTION, edit))


def test_05_the_iteration_bank_columns_are_swapped() -> None:
    def edit(document):
        banks = document["sim_data"]["iteration_records"]["banks"]
        banks["A"], banks["B"] = banks["B"], banks["A"]
    _control("test_19", inspection=_json_mutation(_INSPECTION, edit))


def test_06_the_attempt_result_row_moves() -> None:
    """Every REFUSED / FAILED / SUCCESS assertion in the matrix reads this row."""
    def edit(document):
        document["sim_data"]["run_identity"]["rows"]["last_attempt_result"] = 24
    _control("test_17", inspection=_json_mutation(_INSPECTION, edit))


def test_07_the_simulation_status_row_moves() -> None:
    def edit(document):
        document["sim_data"]["run_identity"]["rows"]["simulation_status"] = 29
    _control("test_17", inspection=_json_mutation(_INSPECTION, edit))


def test_08_the_active_bank_selector_row_moves() -> None:
    def edit(document):
        document["sim_data"]["run_identity"]["rows"]["active_bank"] = 29
    _control("test_17", inspection=_json_mutation(_INSPECTION, edit))


def test_09_the_nonce_counter_row_moves() -> None:
    def edit(document):
        document["sim_data"]["run_identity"]["rows"]["next_auto_nonce"] = 22
    _control("test_17", inspection=_json_mutation(_INSPECTION, edit))


def test_10_the_summary_ladder_rows_shift() -> None:
    def edit(document):
        rows = document["sim_data"]["summary_statistics"]["rows"]
        for key in rows:
            rows[key] = rows[key] + 1
    _control("test_19", inspection=_json_mutation(_INSPECTION, edit))


def test_11_the_contingency_ladder_label_column_moves() -> None:
    def edit(document):
        document["sim_data"]["contingency_ladder"]["label_column"] = "O"
    _control("test_19", inspection=_json_mutation(_INSPECTION, edit))


def test_12_the_iterations_control_defined_name_changes() -> None:
    """A fixture would then write the iteration count into a name production
    does not read, and every run would use whatever was already in the cell."""
    def edit(document):
        document["controls"]["monte_carlo_iterations"]["defined_name"] = "inpIterations"
    _control("test_20", inspection=_json_mutation(_INSPECTION, edit))


def test_13_the_random_seed_control_defined_name_changes() -> None:
    """The seed control is the ONLY thing that selects FIXED versus AUTO."""
    def edit(document):
        document["controls"]["random_seed"]["defined_name"] = "inpSeed"
    _control("test_20", inspection=_json_mutation(_INSPECTION, edit))


def test_14_the_sheet_identity_changes() -> None:
    def edit(document):
        document["sim_data"]["sheet"] = "_SimDataX"
    _control("test_21", inspection=_json_mutation(_INSPECTION, edit))


def test_15_the_required_visibility_is_relaxed() -> None:
    def edit(document):
        document["sim_data"]["required_visibility"] = "hidden"
    _control("test_21", inspection=_json_mutation(_INSPECTION, edit))


def test_16_a_required_key_is_deleted() -> None:
    """A missing key is not a smaller projection; it is a scenario that cannot
    find the cell it exists to read."""
    def edit(document):
        del document["sim_data"]["pending_auto_nonce"]
    _control("test_15", inspection=_json_mutation(_INSPECTION, edit))


def test_17_a_required_identity_row_is_deleted() -> None:
    def edit(document):
        del document["sim_data"]["run_identity"]["rows"]["consumed_auto_nonce"]
    _control("test_17", inspection=_json_mutation(_INSPECTION, edit))


def test_18_an_unapproved_key_is_added() -> None:
    """The allowlist has to refuse the next semantic value too, whatever it is
    called - which is the whole reason it is positive rather than a ban list."""
    def edit(document):
        document["sim_data"]["run_identity"]["tolerance"] = 1e-9
    _control("test_15", inspection=_json_mutation(_INSPECTION, edit))


def test_19_an_unapproved_root_key_is_added() -> None:
    def edit(document):
        document["expected_result_digest"] = "0123456789ABCDEF"
    _control("test_15", inspection=_json_mutation(_INSPECTION, edit))


def test_20_the_projection_gains_a_vocabulary() -> None:
    """Labels are model SEMANTICS, not addresses. The Phase-5 projection had
    them in its first submission and independent review removed them."""
    def edit(document):
        document["sim_data"]["run_identity"]["labels"]["last_attempt_result"] = (
            "SUCCESS / REFUSED / FAILED / AUTO_NONCE_INDETERMINATE"
        )
    _control("test_16", inspection=_json_mutation(_INSPECTION, edit))


def test_21_the_command_surface_loses_an_accessor() -> None:
    def edit(document):
        document["command_surface"]["read_accessors"] = (
            document["command_surface"]["read_accessors"][:5]
        )
    _control("test_22", inspection=_json_mutation(_INSPECTION, edit))


def test_22_the_automation_endpoint_is_renamed() -> None:
    def edit(document):
        document["command_surface"]["automation_endpoint"] = "PCCM_Simulate"
    _control("test_22", inspection=_json_mutation(_INSPECTION, edit))


# ===========================================================================
# B. The PowerShell harness - it must not become a second reader
# ===========================================================================
def test_23_the_harness_hard_codes_the_sidecar_address() -> None:
    """The exact defect the projection exists to prevent."""
    damaged = _swap(
        _PHASE6,
        "    return [string]$Inspection.sim_data.pending_auto_nonce.cell\n",
        "    return 'F21'\n")
    _control("test_05", phase6=damaged)


def test_24_the_harness_hard_codes_the_final_commit_range() -> None:
    damaged = _swap(
        _PHASE6,
        "    $pendingCell = Get-SimPendingCell -Inspection $SimInspection\n",
        "    $pendingCell = Get-SimPendingCell -Inspection $SimInspection\n"
        "    $commitRange = 'D22:D30'\n")
    _control("test_05", phase6=damaged)


def test_25_the_harness_hard_codes_the_sheet_name() -> None:
    damaged = _swap(
        _PHASE6,
        "        $sheet = $sheets.Item([string]$Inspection.sim_data.sheet)\n"
        "        $range = $sheet.Range($Address)\n"
        "        return $range.Value2\n",
        "        $sheet = $sheets.Item('_SimData')\n"
        "        $range = $sheet.Range($Address)\n"
        "        return $range.Value2\n")
    _control("test_06", phase6=damaged)


def test_26_the_harness_hard_codes_a_defined_name() -> None:
    damaged = _swap(
        _PHASE6,
        "            -DefinedName ([string]$controls.monte_carlo_iterations.defined_name) `\n"
        "            -Value ([double]$Iterations)\n",
        "            -DefinedName 'inpMonteCarloIterations' `\n"
        "            -Value ([double]$Iterations)\n")
    _control("test_06", phase6=damaged)


def test_27_the_harness_hard_codes_a_bank_column() -> None:
    damaged = _swap(
        _PHASE6,
        "        return [string]$identity.bank_value_columns.$Bank\n",
        "        if ($Bank -ceq 'A') { return 'D' }\n"
        "        return [string]$identity.bank_value_columns.$Bank\n")
    _control("test_08", phase6=damaged)


def test_28_the_harness_hard_codes_an_identity_row() -> None:
    damaged = _swap(
        _PHASE6,
        "    $row = [int]$identity.rows.$FieldKey\n",
        "    $row = [int]$identity.rows.$FieldKey\n"
        "    if ($FieldKey -ceq 'next_auto_nonce') { $row = 21 }\n")
    _control("test_07", phase6=damaged)


def test_29_the_harness_pastes_an_expected_digest() -> None:
    damaged = _swap(
        _PHASE6,
        "    $expected = $Case.expected_exact\n",
        "    $expected = $Case.expected_exact\n"
        "    $known = '4970DF75235C8F6D'\n")
    _control("test_09", phase6=damaged)


def test_30_the_harness_admits_a_tolerance() -> None:
    """Exactness is the whole comparison policy. A tolerance here would let a
    wrong number pass as a right one."""
    damaged = _swap(
        _PHASE6,
        "    if ($Actual.GetType().FullName -cne 'System.Double') { return $false }\n"
        "    return ([double]$Actual -eq $Expected)\n",
        "    if ($Actual.GetType().FullName -cne 'System.Double') { return $false }\n"
        "    if ([math]::Abs([double]$Actual - $Expected) -lt 1e-9) { return $true }\n"
        "    return ([double]$Actual -eq $Expected)\n")
    _control("test_10", phase6=damaged)


def test_31_the_harness_names_a_failpoint_production_does_not_declare() -> None:
    damaged = _swap(
        _PHASE6,
        "        CandidateBank       = 'Phase6CandidateBank'\n",
        "        CandidateBank       = 'Phase6CandidateWrite'\n")
    _control("test_11", phase6=damaged)


def test_32_the_harness_stops_naming_a_read_accessor() -> None:
    """Six accessors, and P6-ACC must exercise all six. Validating four and
    calling it the surface is the overclaim this control refuses."""
    damaged = _swap(
        _PHASE6,
        "        $currentFingerprint = [string]$Excel.Run('PCCM_CurrentSimulationRequestFingerprint')\n",
        "        $currentFingerprint = $storedFingerprint\n")
    _control("test_12", phase6=damaged)


def test_33_the_harness_invents_a_public_procedure() -> None:
    damaged = _swap(
        _PHASE6,
        "        $storedDigest = [string]$Excel.Run('PCCM_SimulationResultDigest')\n",
        "        $storedDigest = [string]$Excel.Run('PCCM_SimulationDigest')\n")
    _control("test_12", phase6=damaged)


def test_34_the_harness_records_a_skip() -> None:
    """"Not attempted" must be as loud as "failed"."""
    damaged = _swap(
        _PHASE6,
        "            Add-Phase6Result $id 'not attempted' 'FAIL' `\n",
        "            Add-Phase6Result $id 'not attempted' 'SKIP' `\n")
    _control("test_36", phase6=damaged)


def test_35_the_prerequisite_gate_is_removed() -> None:
    damaged = _swap(
        _PHASE6,
        "    if (-not $prerequisiteOk) {\n",
        "    if ($false) {\n")
    _control("test_36", phase6=damaged)


def test_36_the_ledger_violation_stops_being_a_failure() -> None:
    damaged = _swap(
        _PHASE6,
        "    Add-Result 'P6-LDG' 'Phase-6 result ledger: one result per scenario ID' 'FAIL' `\n",
        "    Add-Result 'P6-LDG' 'Phase-6 result ledger: one result per scenario ID' 'PASS' `\n")
    _control("test_37", phase6=damaged)


def test_37_the_recovery_scenarios_stop_restoring() -> None:
    damaged = _swap(
        _PHASE6,
        "            if ($null -ne $fixture) {\n"
        "                $null = Complete-Phase6Fixture -Workbook $Workbook -Inspection $SimInspection `\n"
        "                    -Fixture $fixture -List $list -Label $Id\n"
        "            }\n",
        "")
    _control("test_38", phase6=damaged)


def test_37b_the_unwind_moves_out_of_the_finally_onto_the_success_path() -> None:
    """The submitted defect exactly: a restore that only runs when nothing threw.

    An exception between the fixture write and the restore then leaves the
    modified cell in the workbook, and every later scenario runs against it.
    """
    unwind = (
        "        } finally {\n"
        "            # RESTORATION IS REACHED ON EVERY PATH. An exception between the\n"
        "            # fixture write and the restore is exactly the case that would\n"
        "            # otherwise leave a modified F21 or counter in the workbook and let\n"
        "            # every later scenario run against it.\n"
        "            if ($null -ne $fixture) {\n"
        "                $null = Complete-Phase6Fixture -Workbook $Workbook -Inspection $SimInspection `\n"
        "                    -Fixture $fixture -List $list -Label $Id\n"
        "            }\n"
        "        }\n")
    assert _PHASE6.count(unwind) == 1
    success_path = (
        "            & $Assert $list $before $after $announced $afterSecond $secondAnnounced\n")
    assert _PHASE6.count(success_path) == 1
    damaged = _PHASE6.replace(unwind, "        }\n", 1).replace(
        success_path,
        success_path +
        "            $null = Complete-Phase6Fixture -Workbook $Workbook "
        "-Inspection $SimInspection `\n"
        "                -Fixture $fixture -List $list -Label $Id\n", 1)
    assert damaged != _PHASE6
    _control("test_38", phase6=damaged)


def test_37c_the_fixture_is_only_marked_written_after_a_successful_write() -> None:
    """A raising assignment can still have changed the cell."""
    damaged = _swap(
        _PHASE6,
        "    $Fixture.Written = $true\n"
        "    Set-SimRawCell -Workbook $Workbook -Inspection $Inspection `\n"
        "        -Address $Fixture.Address -Value $Value\n",
        "    Set-SimRawCell -Workbook $Workbook -Inspection $Inspection `\n"
        "        -Address $Fixture.Address -Value $Value\n"
        "    $Fixture.Written = $true\n")
    _control("test_38", phase6=damaged)


def test_37d_a_failed_restoration_stops_latching_contamination() -> None:
    """Continuing produces behavioural evidence from a state the harness made."""
    damaged = _swap(
        _PHASE6,
        "    if (-not $restored) {\n"
        "        Set-Phase6Contaminated -Reason ($Label + ' could not restore ' + $Fixture.Address +\n"
        "            ' (original ' + (Format-SimValue $Fixture.Original) + ')')\n"
        "    }\n",
        "")
    _control("test_38", phase6=damaged)


def test_37e_a_stateful_scenario_runs_on_after_a_failed_restoration() -> None:
    """The guard is what stops the next scenario trusting the workbook."""
    damaged = _swap(
        _PHASE6,
        "        if (-not (Test-Phase6FixtureIntegrity)) {\n"
        "            Add-Phase6Result $Id $Name 'FAIL' `\n",
        "        if ($false) {\n"
        "            Add-Phase6Result $Id $Name 'FAIL' `\n")
    _control("test_38c", phase6=damaged)


def test_37f_the_contamination_flag_stops_latching() -> None:
    damaged = _swap(
        _PHASE6,
        "    if ($script:Phase6FixtureIntegrity) {\n"
        "        $script:Phase6FixtureIntegrity = $false\n",
        "    if ($true) {\n"
        "        $script:Phase6FixtureIntegrity = $true\n")
    _control("test_38c", phase6=damaged)


def test_38_the_restoration_check_stops_reaching_the_checklist() -> None:
    """A restoration failure hidden in a note is fail-open behaviour."""
    damaged = _swap(
        _PHASE6,
        "    return (Add-Check $List ($Label + ': ' + $Fixture.Address + ' is restored exactly') `\n"
        "        (Test-SimSameValue -A $readBack -B $Fixture.Original) `\n"
        "        ('original ' + (Format-SimValue $Fixture.Original) + ', restored ' +\n"
        "         (Format-SimValue $readBack)))\n",
        "    Add-Note ($Label + ': restored ' + (Format-SimValue $readBack))\n"
        "    return $true\n")
    _control("test_38", phase6=damaged)


def test_38b_a_raising_restore_escapes_instead_of_being_recorded() -> None:
    """A cleanup that throws would replace the original scenario failure."""
    damaged = _swap(
        _PHASE6,
        "    } catch {\n"
        "        $failure = Format-Phase6Err $_\n"
        "    }\n"
        "    if (-not [string]::IsNullOrEmpty($failure)) {\n",
        "    }\n"
        "    if ($false) {\n")
    _control("test_38", phase6=damaged)


def test_38c_the_original_scenario_failure_is_discarded_by_the_cleanup() -> None:
    damaged = _swap(
        _PHASE6,
        "        } catch {\n"
        "            # THE ORIGINAL FAILURE IS PRESERVED. Cleanup runs next and must not\n"
        "            # replace this with its own story.\n"
        "            $scenarioFailure = Format-Phase6Err $_\n",
        "        } catch {\n"
        "            $scenarioFailure = ''\n")
    _control("test_38b", phase6=damaged)


def test_39_the_recovery_evidence_is_captured_after_cleanup() -> None:
    """Post-state captured after the restore would describe the harness's own
    tidy-up, not what production left behind."""
    # THE RECOVERY BLOCK'S capture, not RIDMAX's - both carry the marker, so
    # the anchor is the statement that follows only this one.
    evidence_start = "            # POST EVIDENCE IS CAPTURED BEFORE CLEANUP, ALWAYS.\n"
    assert _PHASE6.count(evidence_start) == 2
    end = _PHASE6.index("            & $Assert $list $before $after")
    start = _PHASE6.rindex(evidence_start, 0, end)
    block = _PHASE6[start:end]
    damaged = _PHASE6.replace(block, "", 1).replace(
        "        if (-not [string]::IsNullOrEmpty($scenarioFailure)) {",
        block + "        if (-not [string]::IsNullOrEmpty($scenarioFailure)) {", 1)
    assert damaged != _PHASE6
    _control("test_39", phase6=damaged)


def test_39b_the_run_id_evidence_is_captured_after_cleanup() -> None:
    """The same rule, in the other fixture-writing scenario."""
    start = _PHASE6.index("            # POST EVIDENCE IS CAPTURED BEFORE CLEANUP, ALWAYS.\n"
                          "            $evidence = (Format-Phase6State -State $before -Label 'before') + \"`r`n\" +\n"
                          "                        '    fixture: ' + $runIdCell")
    end = _PHASE6.index("            $null = Add-Check $list 'the endpoint refused' `")
    block = _PHASE6[start:end]
    damaged = _PHASE6.replace(block, "", 1).replace(
        "        if (-not [string]::IsNullOrEmpty($scenarioFailure)) {\n"
        "            $null = Add-Check $list 'P6-RIDMAX: the scenario ran to completion'",
        block +
        "        if (-not [string]::IsNullOrEmpty($scenarioFailure)) {\n"
        "            $null = Add-Check $list 'P6-RIDMAX: the scenario ran to completion'", 1)
    assert damaged != _PHASE6
    _control("test_39", phase6=damaged)


def test_40_the_harness_claims_a_mid_call_observation() -> None:
    damaged = _swap(
        _PHASE6,
        "'    NOTE: this scenario observes only states BEFORE and AFTER a completed ' +",
        "'    NOTE: this scenario reads the workbook while the call is running, ' +")
    _control("test_40", phase6=damaged)


def test_41_the_harness_claims_the_private_consumption_flag_at_runtime() -> None:
    damaged = _swap(
        _PHASE6,
        "'    NOT CLAIMED: the private NonceConsumed projection.' + \"`r`n\" +\n"
        "             '    PowerShell cannot observe it; that remains source evidence.')",
        "'    PROVED HERE: the NonceConsumed projection, read back from the' + \"`r`n\" +\n"
        "             '    published records after every attempt.')")
    _control("test_41", phase6=damaged)


def test_42_the_no_replay_invariant_becomes_digest_inequality() -> None:
    damaged = _swap(
        _PHASE6,
        "                      '. Recorded as evidence; digest inequality is not a contract rule.')",
        "                      '. Two AUTO nonces must always produce two digests.')")
    _control("test_42", phase6=damaged)


def test_43_the_phase6_block_re_executes_the_compile_control() -> None:
    """Reporting a second compile that did not happen is an overclaim."""
    damaged = _swap(
        _PHASE6,
        "        $modules = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })\n"
        "        foreach ($name in @('modSimContract', 'modSimRng', 'modSimSample', 'modSimEngine',\n",
        "        $null = $Workbook.VBProject.VBComponents\n"
        "        $modules = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })\n"
        "        foreach ($name in @('modSimContract', 'modSimRng', 'modSimSample', 'modSimEngine',\n")
    _control("test_43", phase6=damaged)


def test_44_the_phase6_block_starts_its_own_excel() -> None:
    damaged = _swap(
        _PHASE6,
        "    Reset-Phase6ResultLedger\n",
        "    Reset-Phase6ResultLedger\n"
        "    $own = New-Object -ComObject Excel.Application\n")
    _control("test_03", phase6=damaged)


def test_45_the_preflight_stops_checking_the_expectation_corpus() -> None:
    damaged = _swap(
        _PHASE6,
        "            $null = Add-Check $list ($case.id + ' carries ' + $key) `\n"
        "                ($null -ne $case.expected_exact.PSObject.Properties[$key])\n",
        "            $null = Add-Check $list ($case.id + ' carries ' + $key) $true\n")
    damaged = _swap(
        damaged,
        "        foreach ($key in @('result_digest', 'effective_seed', 'iterations_run',\n"
        "                           'summary', 'deterministic_base')) {\n",
        "        foreach ($key in @()) {\n")
    _control("test_44", phase6=damaged)


def test_46_a_required_scenario_is_dropped_from_the_matrix() -> None:
    damaged = _swap(
        _PHASE6,
        "        'P6-RIDMAX', 'P6-AXIS',\n        'P6-SU', 'P6-XX', 'P6-LDG', 'P6-FIN'\n",
        "        'P6-AXIS',\n        'P6-SU', 'P6-XX', 'P6-LDG', 'P6-FIN'\n")
    _control("test_34b", phase6=damaged)


def test_47_the_execution_disclaimer_is_removed() -> None:
    damaged = _swap(
        _PHASE6,
        "    NOTHING HERE HAS BEEN EXECUTED. As submitted, no Windows run has been made.\n",
        "    This harness has been proven on Windows.\n")
    _control("test_45", phase6=damaged)


# ===========================================================================
# C. The parity corpus - the only expected-value authority
# ===========================================================================
def test_48_a_result_digest_is_altered() -> None:
    """The corpus stops being what the oracle emits, so it stops being evidence."""
    def edit(document):
        document["parity_cases"][0]["expected_exact"]["result_digest"] = "0000000000000000"
    _control("test_23", cases=_json_mutation(_CASES, edit))


def test_49_the_golden_analytical_digest_is_altered() -> None:
    def edit(document):
        document["parity_cases"][0]["expected_exact"]["calculation_fingerprint"] = (
            "0123456789ABCDEF"
        )
    _control("test_23", cases=_json_mutation(_CASES, edit))


def test_50_the_iteration_count_drifts_from_the_contract_minimum() -> None:
    def edit(document):
        document["iterations"] = 500
        for case in document["parity_cases"]:
            case["inputs"]["iterations"] = 500
    _control("test_28", cases=_json_mutation(_CASES, edit))


def test_51_a_case_stops_binding_a_plan_case() -> None:
    """Without the binding the two implementations describe "similar" fixtures."""
    def edit(document):
        document["parity_cases"][1]["plan_case_id"] = 99
    _control("test_24", cases=_json_mutation(_CASES, edit))


def test_52_a_case_loses_its_summary_ladder() -> None:
    """A digest match with an unchecked ladder would mean the retained totals
    were right and the statistics layer was never compared."""
    def edit(document):
        del document["parity_cases"][0]["expected_exact"]["summary"]["pv"]
    _control("test_30", cases=_json_mutation(_CASES, edit))


def test_53_a_ladder_loses_a_quantile() -> None:
    def edit(document):
        ladder = document["parity_cases"][0]["expected_exact"]["summary"]["nominal"]
        ladder["quantiles"].pop("P95")
    _control("test_30", cases=_json_mutation(_CASES, edit))


def test_54_the_corpus_admits_a_tolerance() -> None:
    def edit(document):
        document["comparison_policy"] = (
            "EXACT where practical; a relative_tolerance of 1e-12 is admissible."
        )
    _control("test_31", cases=_json_mutation(_CASES, edit))


def test_55_a_second_case_claims_the_golden_binding() -> None:
    """Only plan case 1's analytical identity is independently derivable. A
    second claim would mean a fingerprint had been rebuilt somewhere."""
    def edit(document):
        document["parity_cases"][1]["analytical_identity"][
            "fingerprint_independently_derivable"] = True
    _control("test_27", cases=_json_mutation(_CASES, edit))


def test_56_the_bounds_drift_from_the_contract() -> None:
    def edit(document):
        document["bounds"]["run_id_maximum"] = 2147483646
    _control("test_32", cases=_json_mutation(_CASES, edit))


def test_57_the_vocabulary_order_is_disturbed() -> None:
    """The harness reads `attempt_results[0]` as the never-attempted token and
    `sim_states[0]` as CURRENT, so the order is load-bearing."""
    def edit(document):
        document["vocabulary"]["sim_states"] = ["STALE", "CURRENT", "INVALID"]
    _control("test_32", cases=_json_mutation(_CASES, edit))


def test_58_the_case_count_stops_matching() -> None:
    def edit(document):
        document["case_count"] = 9
    _control("test_23", cases=_json_mutation(_CASES, edit))


def test_59_a_degenerate_fixture_replaces_a_stochastic_one() -> None:
    """Plan case 30's drivers all have min == max, so a parity case built on it
    would pass whatever the RNG did."""
    def edit(document):
        document["parity_cases"][2]["plan_case_id"] = 30
    _control("test_24", cases=_json_mutation(_CASES, edit))


def test_60_the_corpus_is_edited_by_hand_rather_than_generated() -> None:
    """The one mutation that changes NOTHING a scenario compares - and must
    still be refused, because a hand-edited corpus is not oracle output."""
    def edit(document):
        document["purpose"] = document["purpose"] + " Adjusted by hand."
    _control("test_23", cases=_json_mutation(_CASES, edit))


def test_61_the_expectation_corpus_is_emptied() -> None:
    """Fail closed. A corpus with no cases must not read as "nothing to compare"."""
    def edit(document):
        document["parity_cases"] = []
        document["case_count"] = 0
    _control("test_23", cases=_json_mutation(_CASES, edit))


def test_62_the_expectation_corpus_is_corrupted() -> None:
    """Malformed JSON must stop the battery, not be skipped over."""
    _control("test_23", cases="{ this is not json ")


def test_63_the_inspection_projection_is_corrupted() -> None:
    _control("test_14", inspection="{ this is not json ")


def test_64_the_parity_scenario_stops_checking_the_analytical_identity() -> None:
    """Comparing a simulation before proving the workbook holds the model the
    oracle evaluated compares two different questions."""
    damaged = _swap(
        _PHASE6,
        "            if ([bool]$case.analytical_identity.fingerprint_independently_derivable) {\n",
        "            if ($false) {\n")
    _control("test_46", phase6=damaged)


def test_65_the_parity_comparison_runs_before_the_identity_is_established() -> None:
    """Order matters: an identity proved after the comparison proves nothing
    about the comparison that already happened."""
    marker = "            # THE CURRENT ANALYTICAL IDENTITY FIRST, FOR EVERY CASE.\n"
    assert _PHASE6.count(marker) == 1
    start = _PHASE6.index(marker)
    end = _PHASE6.index("            $activeBefore = Get-Phase6ActiveBank `")
    block = _PHASE6[start:end]
    parity = (
        "            Add-Phase6ParityChecks -Workbook $Workbook -Inspection $SimInspection `\n"
        "                -Cases $GateBCases -Case $case -List $list -Bank $target -Label $label\n")
    assert _PHASE6.count(parity) == 1
    damaged = _PHASE6.replace(block, "", 1).replace(parity, parity + block, 1)
    assert damaged != _PHASE6
    _control("test_46", phase6=damaged)


def test_66_the_ladder_comparison_is_dropped_from_the_comparator() -> None:
    damaged = _swap(
        _PHASE6,
        "    foreach ($measure in @('nominal', 'pv')) {\n"
        "        $ladder = $expected.summary.$measure\n",
        "    foreach ($measure in @()) {\n"
        "        $ladder = $expected.summary.$measure\n")
    _control("test_47", phase6=damaged)


def test_67_the_digest_comparison_is_dropped_from_the_comparator() -> None:
    damaged = _swap(
        _PHASE6,
        "        @{ Field = 'result_digest';  Expected = [string]$expected.result_digest },\n",
        "")
    _control("test_47", phase6=damaged)


# ===========================================================================
# D. The corrections of the harness review
# ===========================================================================
def test_68_the_live_prerequisite_demands_the_full_yz_inclusive_phase4_set() -> None:
    """The submitted defect, restored: unsatisfiable by construction.

    Y and Z are post-session lifecycle cases recorded after Excel is torn down,
    and Phase 6 runs inside the live automation session. Demanding them here
    fails P6-PRE before the first simulation, and every stateful scenario with
    it.
    """
    damaged = _swap(
        _PHASE6,
        "        $phase4Prerequisite = @(Get-Phase4PrerequisiteScenarioIds)\n",
        "        $phase4Prerequisite = @(Get-Phase4RequiredScenarioIds)\n")
    _control("test_49", phase6=damaged)


def test_69_the_prerequisite_stops_proving_the_deferral_is_real() -> None:
    """A deferred case that had already run was never a post-session case."""
    damaged = _swap(
        _PHASE6,
        "        $earlyDeferred = @($phase4Deferred | Where-Object { $seen -contains $_ })\n",
        "        $earlyDeferred = @()\n")
    _control("test_49", phase6=damaged)


def test_70_the_prerequisite_demands_the_post_session_phase5_results() -> None:
    """P5-FIN and P5-LDG are recorded after shutdown, like Y and Z."""
    damaged = _swap(
        _PHASE6,
        "        foreach ($id in @('P5-FIN', 'P5-LDG')) {\n"
        "            $null = Add-Check $list `\n"
        "                ('the post-session Phase-5 result ' + $id + ' has not run yet') `\n"
        "                ($seen -notcontains $id)\n"
        "        }\n",
        "        foreach ($id in @('P5-FIN', 'P5-LDG')) {\n"
        "            $null = Add-Check $list `\n"
        "                ('the Phase-5 result ' + $id + ' has run') `\n"
        "                ($seen -contains $id)\n"
        "        }\n")
    _control("test_49", phase6=damaged)


def test_71_the_completeness_verdict_bypasses_the_result_guard() -> None:
    """P6-FIN through Add-Result cannot be seen as a duplicate by the ledger."""
    damaged = _swap(
        _PHASE6,
        "        Add-Phase6Result 'P6-FIN' 'Phase-6 completeness' `\n"
        "            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)\n",
        "        Add-Result 'P6-FIN' 'Phase-6 completeness' `\n"
        "            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)\n")
    _control("test_51", phase6=damaged)


def test_72_the_ledger_verdict_becomes_a_precondition_of_completeness() -> None:
    """Requiring P6-LDG inside the set P6-FIN verifies is the circular ordering."""
    damaged = _swap(
        _PHASE6,
        "        'P6-RIDMAX', 'P6-AXIS'\n    )\n}\n\n# The scenarios that WRITE machine state",
        "        'P6-RIDMAX', 'P6-AXIS', 'P6-LDG'\n    )\n}\n\n# The scenarios that WRITE machine state")
    _control("test_51", phase6=damaged)


def test_73_the_ledger_verdict_is_emitted_before_the_guarded_completeness() -> None:
    """Round 4A: a verdict emitted before the last guarded result is fail-open."""
    damaged = _swap(
        _PHASE6,
        "    try {\n        $list = New-Checklist\n"
        "        $recorded = @($Results | Where-Object { $_.Id -like 'P6-*' })\n",
        "    Add-Phase6LedgerIntegrityResult\n"
        "    try {\n        $list = New-Checklist\n"
        "        $recorded = @($Results | Where-Object { $_.Id -like 'P6-*' })\n")
    _control("test_51", phase6=damaged)


def test_74_the_harness_head_is_reported_as_the_production_baseline() -> None:
    """The two identities the Step-13 authorisation requires to stay distinct."""
    damaged = _swap(
        _PHASE6,
        "function Get-Phase6ProductionBaseline { return '5a5b183' }",
        "function Get-Phase6ProductionBaseline {\n"
        "    return [string](& git rev-parse HEAD 2>$null)\n}")
    _control("test_52", phase6=damaged)


def test_75_the_baseline_pin_drifts_from_the_reviewed_baseline() -> None:
    damaged = _swap(
        _PHASE6,
        "function Get-Phase6ProductionBaseline { return '5a5b183' }",
        "function Get-Phase6ProductionBaseline { return 'd36d5d4' }")
    _control("test_52", phase6=damaged)


def test_76_artefact_identity_passes_without_git() -> None:
    """A runtime result with no attributable revision is weaker evidence, and
    recording "unknown" while passing would pass that weakness off as strength."""
    damaged = _swap(
        _PHASE6,
        "            $null = Add-Check $list `\n"
        "                'the accepted production source can be bound to the baseline' $false `\n",
        "            $null = Add-Check $list `\n"
        "                'the accepted production source can be bound to the baseline' $true `\n")
    _control("test_52", phase6=damaged)


def test_77_the_source_binding_falls_back_to_module_names() -> None:
    """That a project CONTAINS a modSimReport is not whose modSimReport it is."""
    damaged = _swap(
        _PHASE6,
        "                    $accepted = [string](& git -C $RepoRoot rev-parse ($baseline + ':' + $relative) 2>$null)\n",
        "                    $accepted = ''\n")
    _control("test_52", phase6=damaged)


def test_78_the_non_golden_identity_reduces_to_a_note() -> None:
    """A note about earlier evidence is not a check on the current fixture."""
    damaged = _swap(
        _PHASE6,
        "            Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook `\n"
        "                -Inspection $Inspection -Case $planCase -Tolerances $Cases.tolerances `\n"
        "                -Label ($label + ' analytical')\n"
        "            Add-Phase5SuccessStateChecks -List $list -Excel $Excel -Workbook $Workbook `\n"
        "                -Inspection $Inspection -Case $planCase -Cases $Cases `\n"
        "                -Label ($label + ' calc_state')\n",
        "            Add-Note ($label + ': P5-AN already drove this plan case.')\n")
    _control("test_46", phase6=damaged)


def test_79_the_current_fixture_identity_becomes_conditional() -> None:
    """Running the comparators only for the golden case is the same gap."""
    damaged = _swap(
        _PHASE6,
        "            Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook `\n",
        "            if ([bool]$case.analytical_identity.fingerprint_independently_derivable) {\n"
        "            Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook `\n")
    damaged = _swap(
        damaged,
        "                -Label ($label + ' calc_state')\n",
        "                -Label ($label + ' calc_state')\n            }\n")
    _control("test_46", phase6=damaged)


def test_80_the_stored_fingerprint_stops_being_proved_current() -> None:
    """A stale analytical digest names a model that has since moved."""
    damaged = _swap(
        _PHASE6,
        "                ((-not [string]::IsNullOrEmpty($calcCurrent)) -and ($calcStored -ceq $calcCurrent)) `\n",
        "                ($true) `\n")
    _control("test_46", phase6=damaged)


def test_81_the_final_log_omits_the_phase6_summary_identity() -> None:
    """A Step-13 run that finished by claiming only a Phase-4/Phase-5 test had
    run would understate what the log is evidence of."""
    damaged = _swap(
        _HARNESS,
        "    Write-Host 'PHASE-4 / PHASE-5 / PHASE-6 FUNCTIONAL TEST: ALL CHECKS PASSED' "
        "-ForegroundColor Green\n",
        "    Write-Host 'PHASE-4 / PHASE-5 FUNCTIONAL TEST: ALL CHECKS PASSED' "
        "-ForegroundColor Green\n")
    _control("test_02", harness=damaged)


def test_82_the_final_log_drops_the_phase6_scenario_count() -> None:
    start = _HARNESS.index('# NAMED, NOT INFERRED.')
    end = _HARNESS.index("Write-Host ''\nif ($failed.Count -eq 0) {")
    block = _HARNESS[start:end]
    damaged = _HARNESS.replace(block, "", 1)
    assert damaged != _HARNESS
    _control("test_02", harness=damaged)


def test_83_the_driver_stops_emitting_the_phase6_ledger_verdict() -> None:
    """Without it, a duplicate P6-FIN attempt is a Note and the run finishes green."""
    damaged = _swap(_HARNESS, "Add-Phase6LedgerIntegrityResult\n", "")
    _control("test_51", harness=damaged)


def test_84_the_driver_emits_the_phase6_ledger_before_the_scenarios() -> None:
    damaged = _swap(
        _HARNESS,
        "        Invoke-Phase6GateBScenarios -Excel $excel -Workbook $wb -Manifest $manifest `\n",
        "        Add-Phase6LedgerIntegrityResult\n"
        "        Invoke-Phase6GateBScenarios -Excel $excel -Workbook $wb -Manifest $manifest `\n")
    _control("test_51", harness=damaged)


def test_85_the_driver_substitutes_a_placeholder_for_a_missing_commit() -> None:
    """"unknown" and a PASS is a weakness passed off as strength."""
    damaged = _swap(
        _HARNESS,
        "    $harnessCommit = ''\n",
        "    $harnessCommit = 'unknown (git was not available on this machine)'\n")
    _control("test_52", harness=damaged)


# ===========================================================================
# E. The final pre-execution corrections
# ===========================================================================
# THE STATE EACH OF THE FIRST THREE MODELS. Not a hypothetical: it is the exact
# intermediate the accepted Phase-5 guard produces.
#
#   P5-X recorded once, PASS, visible in $Results
#   a second P5-X attempt refused and recorded in Phase5LedgerViolations
#   no P5 result anywhere reads FAIL
#   P5-LDG deferred until after Phase 6, so no verdict exists yet
#
# A P6-PRE that scans recorded results alone sees a clean Phase 5 and runs the
# whole Step-13 matrix on top of a known harness-integrity violation.
def test_86_the_prerequisite_stops_reading_the_pending_phase5_ledger() -> None:
    damaged = _swap(
        _PHASE6,
        "        $phase5LedgerViolations = @(Get-Phase5LedgerViolations)\n",
        "        $phase5LedgerViolations = @()\n")
    _control("test_54", phase6=damaged)


def test_87_the_pending_violation_count_stops_being_required_to_be_zero() -> None:
    damaged = _swap(
        _PHASE6,
        "            ($phase5LedgerViolations.Count -eq 0) `\n",
        "            ($phase5LedgerViolations.Count -ge 0) `\n")
    _control("test_54", phase6=damaged)


def test_88_the_ledger_check_is_dropped_entirely() -> None:
    start = _PHASE6.index(
        "        # THE PENDING LEDGER STATE, READ DIRECTLY.")
    end = _PHASE6.index("        $prerequisiteOk = Test-ChecklistOk $list")
    damaged = _PHASE6.replace(_PHASE6[start:end], "", 1)
    assert damaged != _PHASE6
    _control("test_54", phase6=damaged)


def test_89_the_phase6_block_emits_the_phase5_ledger_verdict_early() -> None:
    """Converting the violation into a FAIL early would answer the question, but
    it moves an accepted Phase-5 result out of its settled lifecycle position."""
    damaged = _swap(
        _PHASE6,
        "        $phase5LedgerViolations = @(Get-Phase5LedgerViolations)\n",
        "        Add-Phase5LedgerIntegrityResult\n"
        "        $phase5LedgerViolations = @(Get-Phase5LedgerViolations)\n")
    _control("test_54", phase6=damaged)


def test_90_the_recovery_fixture_verdict_is_discarded() -> None:
    """A failed establishment check followed by a real simulation is behavioural
    evidence against a fixture the harness proved was not established."""
    damaged = _swap(
        _PHASE6,
        "            $fixtureOk = Set-Phase6CellFixture -Workbook $Workbook -Inspection $SimInspection `\n"
        "                -Fixture $fixture -Value $FixtureValue -List $list -Label $Id\n",
        "            $null = Set-Phase6CellFixture -Workbook $Workbook -Inspection $SimInspection `\n"
        "                -Fixture $fixture -Value $FixtureValue -List $list -Label $Id\n")
    _control("test_55", phase6=damaged)


def test_91_the_recovery_fixture_gate_stops_stopping_anything() -> None:
    """The verdict is captured and then ignored - the subtler form of the same
    defect, and the one a "does the check exist" detector would miss."""
    damaged = _swap(
        _PHASE6,
        "            if (-not $fixtureOk) {\n"
        "                throw ('the direct machine-state fixture at ' + $Address +\n"
        "                       ' could not be established exactly; production was not invoked')\n"
        "            }\n",
        "            if (-not $fixtureOk) {\n"
        "                Add-Note ('the fixture at ' + $Address + ' did not verify')\n"
        "            }\n")
    _control("test_55", phase6=damaged)


def test_92_the_run_id_fixture_verdict_is_discarded() -> None:
    damaged = _swap(
        _PHASE6,
        "            $fixtureOk = Set-Phase6CellFixture -Workbook $Workbook -Inspection $SimInspection `\n"
        "                -Fixture $fixture -Value ([double]$bounds.run_id_maximum) -List $list `\n"
        "                -Label 'P6-RIDMAX'\n",
        "            $null = Set-Phase6CellFixture -Workbook $Workbook -Inspection $SimInspection `\n"
        "                -Fixture $fixture -Value ([double]$bounds.run_id_maximum) -List $list `\n"
        "                -Label 'P6-RIDMAX'\n")
    _control("test_55", phase6=damaged)


def test_93_the_fixture_gate_moves_after_the_simulation() -> None:
    """A gate downstream of the invocation gates nothing."""
    gate = (
        "            if (-not $fixtureOk) {\n"
        "                throw ('the direct machine-state fixture at ' + $Address +\n"
        "                       ' could not be established exactly; production was not invoked')\n"
        "            }\n")
    assert _PHASE6.count(gate) == 1
    after = "            $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection\n"
    damaged = _PHASE6.replace(gate, "", 1).replace(after, after + gate, 1)
    assert damaged != _PHASE6
    _control("test_55", phase6=damaged)


def test_94_the_whole_tree_freeze_returns_to_the_pccm_working_directory() -> None:
    """The submitted shape: `-C <repo>/pccm` with repository-root pathspecs,
    which match nothing and make `--quiet` exit 0 whatever the tree holds."""
    damaged = _swap(
        _PHASE6,
        "        [string]$HarnessCommit, [string]$RepoRoot\n",
        "        [string]$HarnessCommit, [string]$PccmRoot\n")
    damaged = damaged.replace("$RepoRoot", "$PccmRoot")
    _control("test_56", phase6=damaged)


def test_95_only_the_freeze_command_reverts_to_the_wrong_root() -> None:
    """The narrow form: the blob checks stay correct and only the whole-tree
    statement is aimed at the wrong directory."""
    damaged = _swap(
        _PHASE6,
        "            $null = & git -C $RepoRoot diff --quiet $baseline -- 'pccm/src' 'pccm/spec' 2>$null\n",
        "            $null = & git -C ($RepoRoot + '/pccm') diff --quiet $baseline -- 'pccm/src' 'pccm/spec' 2>$null\n")
    _control("test_56", phase6=damaged)


def test_96_the_freeze_pathspec_coverage_check_is_dropped() -> None:
    """A freeze proved by a pathspec that names nothing is not a freeze."""
    damaged = _swap(
        _PHASE6,
        "            $tracked = @(& git -C $RepoRoot ls-tree -r --name-only $baseline -- 'pccm/src' 'pccm/spec' 2>$null)\n"
        "            $null = Add-Check $list `\n"
        "                'the freeze pathspec matches the production trees it names' `\n"
        "                ($tracked.Count -gt 0) ('files under the pathspec at the baseline: ' + $tracked.Count)\n",
        "")
    _control("test_56", phase6=damaged)


def test_97_the_driver_passes_the_pccm_subtree_as_the_repository_root() -> None:
    damaged = _swap(
        _HARNESS,
        "            -Results $results -HarnessCommit $harnessCommit -RepoRoot $repoRoot\n",
        "            -Results $results -HarnessCommit $harnessCommit -RepoRoot $pccmRoot\n")
    _control("test_56", harness=damaged)


# ===========================================================================
# F. The Windows PowerShell 5.1 parse defect that ended Step-13 Run 1
# ===========================================================================
# The projection carried a JSON object whose property NAME was the empty string
# - the contract's legitimate "no bank has ever been published" selector key -
# and Windows PowerShell 5.1's ConvertFrom-Json threw PSArgumentException on the
# whole artefact. The preflight aborted before Excel was started.
#
# The four controls below restore each way back into that failure.
def test_98_the_selector_map_returns_to_an_object_with_a_blank_key() -> None:
    """The submitted shape, exactly: a JSON property whose name is ''."""
    def edit(document):
        document["publication"]["candidate_target"] = {
            ("" if entry["active_bank"] is None else entry["active_bank"]):
                entry["candidate_bank"]
            for entry in document["publication"]["candidate_target"]
        }
    _control("test_59", inspection=_json_mutation(_INSPECTION, edit))


def test_99_any_other_empty_object_key_is_refused() -> None:
    """`candidate_target` is where it happened, not the only place it could."""
    def edit(document):
        document["sim_data"]["run_identity"]["labels"][""] = "unnamed"
    _control("test_59", inspection=_json_mutation(_INSPECTION, edit))


def test_100_the_blank_selector_entry_is_dropped() -> None:
    """A workbook that has never published would then have no candidate target,
    and the very first run of a clean workbook is exactly that state."""
    def edit(document):
        document["publication"]["candidate_target"] = [
            entry for entry in document["publication"]["candidate_target"]
            if entry["active_bank"] is not None
        ]
    _control("test_61", inspection=_json_mutation(_INSPECTION, edit))


def test_101_the_blank_selector_entry_is_duplicated() -> None:
    """Two answers is not an answer; the selector must fail closed on it."""
    def edit(document):
        entries = document["publication"]["candidate_target"]
        blank = [entry for entry in entries if entry["active_bank"] is None][0]
        entries.append(dict(blank))
    _control("test_61", inspection=_json_mutation(_INSPECTION, edit))


def test_102_a_sentinel_replaces_the_blank_key() -> None:
    """A magic "BLANK" token would be a second semantic authority."""
    def edit(document):
        for entry in document["publication"]["candidate_target"]:
            if entry["active_bank"] is None:
                entry["active_bank"] = "BLANK"
    _control("test_61", inspection=_json_mutation(_INSPECTION, edit))


def test_103_a_selector_mapping_is_lost_in_the_reshape() -> None:
    def edit(document):
        document["publication"]["candidate_target"] = (
            document["publication"]["candidate_target"][:2])
    _control("test_61", inspection=_json_mutation(_INSPECTION, edit))


def test_104_a_selector_entry_gains_an_unapproved_key() -> None:
    def edit(document):
        document["publication"]["candidate_target"][0]["note"] = "first run"
    _control("test_61", inspection=_json_mutation(_INSPECTION, edit))


def test_105_the_harness_hard_codes_the_blank_selector_answer() -> None:
    """Bypassing the projection puts the A/B rule back in PowerShell."""
    damaged = _swap(
        _PHASE6,
        "    $entries = @($Inspection.publication.candidate_target)\n",
        "    if ([string]::IsNullOrEmpty($ActiveBank)) { return 'A' }\n"
        "    $entries = @($Inspection.publication.candidate_target)\n")
    _control("test_62", phase6=damaged)


def test_106_the_selector_returns_to_an_empty_property_lookup() -> None:
    """The regression the Windows host cannot survive."""
    damaged = _swap(
        _PHASE6,
        "    $matched = @($entries | Where-Object {\n",
        "    $map = $Inspection.publication.candidate_target\n"
        "    $key = $ActiveBank\n"
        "    if ($null -ne $map.PSObject.Properties[$key]) { return [string]$map.$key }\n"
        "    $matched = @($entries | Where-Object {\n")
    _control("test_62", phase6=damaged)


def test_107_the_selector_stops_failing_closed_on_a_duplicate() -> None:
    damaged = _swap(
        _PHASE6,
        "    if ($matched.Count -gt 1) {\n",
        "    if ($false) {\n")
    _control("test_62", phase6=damaged)


def test_108_the_selector_stops_failing_closed_on_no_match() -> None:
    damaged = _swap(
        _PHASE6,
        "    if ($matched.Count -eq 0) {\n"
        "        throw ('the publication candidate_target projection has no entry for active bank ' +\n"
        "               [char]39 + $ActiveBank + [char]39)\n"
        "    }\n",
        "")
    _control("test_62", phase6=damaged)


def test_109_the_documented_command_would_overwrite_run_1s_log() -> None:
    """An aborted attempt is evidence. Run 1 is what found the ConvertFrom-Json
    defect, and a command that redirects over its log destroys the record."""
    doc_path = conformance.PCCM_ROOT / "docs" / "phase6_step13_gate_b.md"
    original = doc_path.read_text(encoding="utf-8")
    damaged = original.replace(
        "  *> .\\pccm\\bootstrap\\windows\\phase6_gate_b_run4.log",
        "  *> .\\pccm\\bootstrap\\windows\\phase6_gate_b_run3.log", 1)
    assert damaged != original
    saved = conformance.PCCM_ROOT
    import tempfile
    with tempfile.TemporaryDirectory(prefix="pccm-step13-doc-") as name:
        root = Path(name)
        (root / "docs").mkdir()
        (root / "docs" / "phase6_step13_gate_b.md").write_text(
            damaged, encoding="utf-8")
        conformance.PCCM_ROOT = root
        try:
            refused = False
            try:
                conformance.test_63_the_documented_log_name_does_not_overwrite_a_recorded_run()
            except AssertionError as error:
                refused = True
                assert "overwritten" in str(error), error
        finally:
            conformance.PCCM_ROOT = saved
    assert refused, "the mutation survived: the run ledger control is vacuous"


# ===========================================================================
# G. The Run-2 harness finalisation defect
# ===========================================================================
def test_110_the_ledger_verdict_reads_an_unwrapped_collection() -> None:
    """The submitted shape, and it threw on every clean run.

    Zero violations is the NORMAL case; a function returning an empty collection
    emits zero pipeline objects, so the assignment lands $null and StrictMode
    turns `.Count` into a hard error.
    """
    damaged = _swap(
        _PHASE6,
        "    $violations = @(Get-Phase6LedgerViolations)\n",
        "    $violations = Get-Phase6LedgerViolations\n")
    _control("test_64", phase6=damaged)


def test_111_another_collection_helper_loses_its_caller_side_wrapper() -> None:
    """The rule is the class, not the one call site that failed."""
    damaged = _swap(
        _PHASE6,
        "    $declared = @(Get-Phase6ScenarioIds)\n",
        "    $declared = Get-Phase6ScenarioIds\n")
    _control("test_64", phase6=damaged)


def test_112_the_fail_arm_reports_only_the_first_violation() -> None:
    """A run that attempted three duplicate results has three facts to answer for."""
    damaged = _swap(
        _PHASE6,
        "         ($violations -join ' | '))\n",
        "         $violations[0])\n")
    _control("test_64", phase6=damaged)


def test_113_the_zero_violation_branch_can_reach_a_failure() -> None:
    damaged = _swap(
        _PHASE6,
        "        Add-Result 'P6-LDG' 'Phase-6 result ledger: one result per scenario ID' 'PASS' `\n",
        "        Add-Result 'P6-LDG' 'Phase-6 result ledger: one result per scenario ID' 'FAIL' `\n")
    _control("test_64", phase6=damaged)


def test_114_the_ledger_verdict_loses_its_emitted_once_flag() -> None:
    """Many duplicate attempts must still produce exactly one P6-LDG."""
    damaged = _swap(
        _PHASE6,
        "    if ($script:Phase6LedgerReported) { return }\n"
        "    $script:Phase6LedgerReported = $true\n"
        "    $violations = @(Get-Phase6LedgerViolations)\n",
        "    $violations = @(Get-Phase6LedgerViolations)\n")
    _control("test_51", phase6=damaged)


def test_115_a_phase6_result_is_recorded_after_the_ledger_verdict() -> None:
    """A result the verdict could not see is a result the ledger did not check."""
    damaged = _swap(
        _HARNESS,
        "Add-Phase6LedgerIntegrityResult\n",
        "Add-Phase6LedgerIntegrityResult\nAdd-Phase6Result 'P6-XX' 'late' 'PASS'\n")
    _control("test_65", harness=damaged)


def test_116_the_production_baseline_pin_lags_the_repaired_source() -> None:
    """After a production repair the pin must move with it, or every runtime
    result is attributed to source that no longer exists."""
    damaged = _swap(
        _PHASE6,
        "function Get-Phase6ProductionBaseline { return '5a5b183' }",
        "function Get-Phase6ProductionBaseline { return 'bc7949b' }")
    _control("test_52", phase6=damaged)


def _phase5_freeze_refuses(damaged: str) -> str:
    """Run the Phase-5 freeze control against damaged CURRENT content.

    Only the current side is swapped. The accepted side is read from git by
    path, so pointing the module at a temporary file would break that lookup
    rather than test the control.
    """
    saved = conformance._current_lines

    def current(path):
        if path == conformance.PHASE5:
            return damaged.splitlines()
        return saved(path)

    conformance._current_lines = current
    try:
        conformance.test_01b_the_only_change_to_the_phase5_scenarios_is_the_scoped_grant()
    except AssertionError as error:
        return str(error)
    finally:
        conformance._current_lines = saved
    raise AssertionError("the mutation survived the Phase-5 freeze control")


def test_117_the_phase5_block_loses_something_beyond_the_stale_assertion() -> None:
    """The Phase-5 correction is one assertion. Anything else is a rewrite."""
    original = conformance.PHASE5.read_text(encoding="utf-8")
    damaged = original.replace(
        "    $null = Add-Check $list 'the Phase-4 matrix has 0 FAIL' ($failed.Count -eq 0) `\n",
        "", 1)
    assert damaged != original
    message = _phase5_freeze_refuses(damaged)
    assert "beyond the stale assertion" in message, message


def test_118_the_scoped_grant_correction_is_reverted() -> None:
    """The stale assertion restored, in the file the Step-13 freeze watches."""
    original = conformance.PHASE5.read_text(encoding="utf-8")
    damaged = original.replace(
        "$null = Add-Check $list 'RunSimulation is permitted in modSimReport and nowhere else' `",
        "$null = Add-Check $list 'RunSimulation is still forbidden in every module' `", 1)
    assert damaged != original
    _phase5_freeze_refuses(damaged)
