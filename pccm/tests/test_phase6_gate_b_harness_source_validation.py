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
_ORACLE = conformance.GATE_B_ORACLE_PATH.read_text(encoding="utf-8")

# The pinned production baseline, read from the conformance module rather
# than spelled again here: a repair that moves the baseline must not leave
# three stale mutation anchors behind, silently matching nothing.
_BASELINE_PIN = ("function Get-Phase6ProductionBaseline { return '"
                 + conformance.PRODUCTION_BASELINE + "' }")


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
               phase6: str | None = None, harness: str | None = None,
               oracle: str | None = None):
    saved = (conformance.INSPECTION_PATH, conformance.GATE_B_CASES_PATH,
             conformance.PHASE6, dict(conformance._CACHE), conformance.HARNESS,
             conformance.GATE_B_ORACLE_PATH)
    with tempfile.TemporaryDirectory(prefix="pccm-step13-mutation-") as name:
        temp = Path(name)
        conformance._CACHE.clear()
        try:
            for damaged, original, attribute in (
                (inspection, _INSPECTION, "INSPECTION_PATH"),
                (cases, _CASES, "GATE_B_CASES_PATH"),
                (phase6, _PHASE6, "PHASE6"),
                (harness, _HARNESS, "HARNESS"),
                (oracle, _ORACLE, "GATE_B_ORACLE_PATH"),
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
            conformance.GATE_B_ORACLE_PATH = saved[5]
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
        "        foreach ($key in @('effective_seed', 'iterations_run',\n"
        "                           'rng_version', 'sim_method_version')) {\n",
        "        foreach ($key in @()) {\n")
    damaged = _swap(
        damaged,
        "        foreach ($withdrawn in @('result_digest', 'summary', 'deterministic_base')) {\n",
        "        foreach ($withdrawn in @()) {\n")
    _control("test_44", phase6=damaged)


def test_46_a_required_scenario_is_dropped_from_the_matrix() -> None:
    damaged = _swap(
        _PHASE6,
        "        'P6-RIDMAX', 'P6-AXIS',\n        'P6-SU', 'P6-XX', 'P6-LDG', 'P6-FIN'\n",
        "        'P6-AXIS',\n        'P6-SU', 'P6-XX', 'P6-LDG', 'P6-FIN'\n")
    _control("test_34b", phase6=damaged)


def test_47_the_execution_boundary_is_removed_from_the_banner() -> None:
    """The banner says which runs have executed AND which corrections they did
    not exercise. Dropping the second half turns a record into a claim."""
    damaged = _swap(
        _PHASE6,
        "    AND THE BOUNDARY SURVIVES THE PASS. An all-green run proves the scenarios\n"
        "    that ran, not the arms this harness cannot reach. The genuine\n",
        "    EVERYTHING IS NOW PROVEN, including the arms this harness cannot reach.\n"
        "    Formerly: the genuine\n")
    damaged = damaged.replace("were NOT induced and are not claimed - they remain static-only, and §5 of",
                              "are covered. See §5 of")
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
        del document["measurements"][0]["summary"]["pv"]
    _control("test_30", oracle=_json_mutation(_ORACLE, edit))


def test_53_a_ladder_loses_a_quantile() -> None:
    def edit(document):
        document["measurements"][0]["summary"]["nominal"]["quantiles"].pop("P95")
    _control("test_30", oracle=_json_mutation(_ORACLE, edit))


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
        "                -Cases $GateBCases -Case $case -List $list -Bank $target -Label $label `\n"
        "                -Measured $measured[0]\n")
    assert _PHASE6.count(parity) == 1
    damaged = _PHASE6.replace(block, "", 1).replace(parity, parity + block, 1)
    assert damaged != _PHASE6
    _control("test_46", phase6=damaged)


def test_66_the_ladder_comparison_is_dropped_from_the_comparator() -> None:
    damaged = _swap(
        _PHASE6,
        "    foreach ($measure in @('nominal', 'pv')) {\n"
        "        $ladder = $Measured.summary.$measure\n",
        "    foreach ($measure in @()) {\n"
        "        $ladder = $Measured.summary.$measure\n")
    _control("test_47", phase6=damaged)


def test_67_the_request_fingerprint_comparison_is_dropped_from_the_comparator() -> None:
    """The digest is no longer compared here - Step 0 §10.4 never promised it
    across languages - but the request fingerprint is an EXACT identity field and
    dropping it would let a run of the wrong request compare as parity."""
    damaged = _swap(
        _PHASE6,
        "        @{ Field = 'request_fingerprint'; Expected = $(\n",
        "        @{ Field = 'withdrawn_fingerprint'; Expected = $(\n")
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
        _BASELINE_PIN,
        "function Get-Phase6ProductionBaseline {\n"
        "    return [string](& git rev-parse HEAD 2>$null)\n}")
    _control("test_52", phase6=damaged)


def test_75_the_baseline_pin_drifts_from_the_reviewed_baseline() -> None:
    damaged = _swap(
        _PHASE6,
        _BASELINE_PIN,
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
        "$harnessCommit = ''\n",
        "$harnessCommit = 'unknown (git was not available on this machine)'\n")
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
        "        [string]$HarnessCommit, [string]$RepoRoot, $ArtefactIdentity,\n",
        "        [string]$HarnessCommit, [string]$PccmRoot, $ArtefactIdentity,\n")
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
        "            -Results $results -HarnessCommit $harnessCommit -RepoRoot $repoRoot `\n",
        "            -Results $results -HarnessCommit $harnessCommit -RepoRoot $pccmRoot `\n")
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
    # The log the NEXT authorised run writes, redirected onto the last one that
    # already happened. The numbers come from the ledger, not from this file.
    import re as _re
    recorded = sorted(int(n) for n in _re.findall(r"\|\s*\*\*Run (\d+)\*\*\s*\|", original))
    nxt, prev = max(recorded), max(recorded) - 1
    damaged = original.replace(
        f"  *> .\\pccm\\bootstrap\\windows\\phase6_gate_b_run{nxt}.log",
        f"  *> .\\pccm\\bootstrap\\windows\\phase6_gate_b_run{prev}.log", 1)
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
        _BASELINE_PIN,
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
    assert "moved from" in message, message


def test_118_the_scoped_grant_correction_is_reverted() -> None:
    """The stale assertion restored, in the file the Step-13 freeze watches."""
    original = conformance.PHASE5.read_text(encoding="utf-8")
    damaged = original.replace(
        "$null = Add-Check $list 'RunSimulation is permitted in modSimReport and nowhere else' `",
        "$null = Add-Check $list 'RunSimulation is still forbidden in every module' `", 1)
    assert damaged != original
    _phase5_freeze_refuses(damaged)


# ===========================================================================
# H. The active preamble
# ===========================================================================
_STEP13_DOC = conformance.PCCM_ROOT / "docs" / "phase6_step13_gate_b.md"

# The preamble as it stood at 7aa1ef3, after three Windows runs had happened.
_STALE_PREAMBLE = """# PCCM Phase 6 — Step 13: the Windows/Excel Gate-B runtime harness

**Status: HARNESS SOURCE, submitted for review. NO STEP-13 SCENARIO HAS EXECUTED.**

Three Windows runs have been attempted. Run 1 aborted in the preflight before
Excel started. Run 2 reached Excel and was stopped at the compile prerequisite
by a genuine production compile defect. Run 3 **proved that repair on the real
VBA compiler** and completed the Phase-4 matrix 35/35, but one stale Phase-5
harness assertion failed, and the Phase-6 matrix correctly failed closed behind
it. See the run ledger in [§8](#8-windows-run-ledger).

**Not one Phase-6 procedure has executed**, no simulation has been performed and
no parity comparison has been made. Every claim in this document about Step-13
behaviour remains a statement about source.

```
static / source evidence   !=   Windows / Excel runtime evidence
```

That line is the whole point of Step 13, and this document keeps it in front of
the reader rather than at the end.

---
"""


def _damage_preamble(rewrite) -> str:
    """Return the real document with ONLY its preamble rewritten.

    Section 8 is left byte-identical on purpose: these mutations must prove the
    control reads the CURRENT claim, not that it can find the word `executed`
    somewhere in the run history.
    """
    original = _STEP13_DOC.read_text(encoding="utf-8")
    cut = original.index("\n## ")
    damaged = rewrite(original[:cut]) + original[cut:]
    assert damaged != original, "the mutation changed nothing"
    return damaged


def _preamble_control_refuses(damaged: str) -> str:
    """Run the active-preamble control over a damaged copy of the document."""
    saved = conformance.PCCM_ROOT
    with tempfile.TemporaryDirectory(prefix="pccm-step13-preamble-") as name:
        root = Path(name)
        (root / "docs").mkdir()
        (root / "docs" / "phase6_step13_gate_b.md").write_text(
            damaged, encoding="utf-8")
        conformance.PCCM_ROOT = root
        try:
            conformance.test_66_the_active_preamble_states_what_has_run_and_what_has_not()
        except AssertionError as error:
            return str(error)
        finally:
            conformance.PCCM_ROOT = saved
    raise AssertionError("the mutation survived the active-preamble control")


def test_119_the_stale_preamble_is_restored() -> None:
    """The submitted wording: NO STEP-13 SCENARIO HAS EXECUTED, sitting above a
    ledger recording three attempts and two runtime-proven closures."""
    damaged = _damage_preamble(lambda _: _STALE_PREAMBLE)
    message = _preamble_control_refuses(damaged)
    assert "stale claim" in message, message
    # AND THE HISTORY IT SITS ABOVE IS UNTOUCHED: the control refused the
    # present tense, not the record of what Run 2 established.
    assert "Neither behavioural matrix executed." in damaged


# The preamble as it stood at 6cb7f06 - accurate for Runs 1-3, and falsified by
# Run 4 executing the behavioural matrix.
_RUNS_1_TO_3_PREAMBLE = '# PCCM Phase 6 — Step 13: the Windows/Excel Gate-B runtime harness\n\n**Status: HARNESS SOURCE UNDER RUNTIME VALIDATION. Runs 1–3 have executed; the\nPhase-6 behavioural matrix `P6-ART` through `P6-AXIS` has not yet executed.**\n\nStep-13 runtime has **partially** executed, through Runs 1–3. Run 1 aborted in\nthe preflight before Excel started. Run 2 reached Excel and was stopped at the\ncompile prerequisite by a genuine production compile defect. Run 3 **proved that\nrepair on the real VBA compiler**, completed the Phase-4 matrix 35/35, and\nreached finalisation with `P6-LDG` PASS; one stale Phase-5 harness assertion\nfailed, and the Phase-6 matrix correctly failed closed behind it. See the run\nledger in [§8](#8-windows-run-ledger).\n\n**Runtime-proven, and no further.** Windows evidence now carries exactly these\nclaims: the accepted workbook builds, opens and **compiles** on the real VBA\ncompiler; the Phase-4 lifecycle matrix runs 35/35 with a natural shutdown and a\nclean COM release ledger; 38 of the 39 Phase-5 Gate-B scenarios pass; and the\nPhase-6 result ledger finalises. Those claims, and only those, have moved from\nsource evidence to runtime evidence.\n\n**Still source-only.** The Phase-6 behavioural matrix, `P6-ART` through\n`P6-AXIS`, remains **unexecuted**. No production Phase-6 simulation procedure has\nexecuted, no simulation has been performed, and no oracle parity result has been\nestablished. Every claim in this document about Phase-6 *behaviour* remains a\nstatement about source.\n\n```\nstatic / source evidence   !=   Windows / Excel runtime evidence\n```\n\nThat line is the whole point of Step 13, and this document keeps it in front of\nthe reader rather than at the end — with the boundary drawn where Runs 1–3\nactually left it, and not where it stood before the first run.\n\n---\n'


def test_120_the_preamble_describes_the_run_before_last() -> None:
    """The wording ACCEPTED for Runs 1-3, restored after Run 4 executed the
    matrix. It was true when written; the ledger below it is what makes it
    false, which is exactly what this control has to notice."""
    damaged = _damage_preamble(lambda _: _RUNS_1_TO_3_PREAMBLE)
    message = _preamble_control_refuses(damaged)
    assert "the ledger records Run 4" in message, message
    assert "Neither behavioural matrix executed." in damaged


def test_120b_the_unexecuted_claim_survives_the_run_that_executed_it() -> None:
    """The surgical form: everything else in the preamble is current, and only
    the claim Run 4 falsified is put back."""
    damaged = _damage_preamble(
        lambda head: head.replace(
            "```\nstatic / source evidence",
            "The Phase-6 behavioural matrix, `P6-ART` through `P6-AXIS`, remains\n"
            "unexecuted.\n\n```\nstatic / source evidence", 1))
    message = _preamble_control_refuses(damaged)
    assert "says the Phase-6 matrix is unexecuted" in message, message


def test_121_the_source_only_claim_is_swept_back_across_step_13() -> None:
    """`Every claim ... about Step-13 behaviour` denies the compile, lifecycle,
    ledger and behavioural closures the runs actually established."""
    damaged = _damage_preamble(
        lambda head: head.replace(
            "```\nstatic / source evidence",
            "Every claim in this document about Step-13 behaviour remains a\n"
            "statement about source.\n\n```\nstatic / source evidence", 1))
    message = _preamble_control_refuses(damaged)
    assert "every Step-13 claim" in message, message


def test_122_the_closing_run_is_reported_without_what_it_reported() -> None:
    """A matrix that ran is not a matrix that passed, and while anything was open
    the preamble had to name it. Once the ledger records an all-green run the
    demand changes rather than disappearing: the preamble must say what that run
    ACTUALLY reported, because "closed" without a tally is the same unsupported
    claim in the other direction."""
    def rewrite(head):
        for phrase in ("103/0/0", "29/29", "all green"):
            head = head.replace(phrase, "as expected")
        return head
    damaged = _damage_preamble(rewrite)
    message = _preamble_control_refuses(damaged)
    assert "does not state what Run" in message, message


def test_123_the_runtime_proven_claims_lose_their_bound() -> None:
    """Naming what Windows proved is half of it. An unbounded list invites the
    reader to assume the rest followed."""
    def rewrite(head):
        # Every phrasing of the bound at once. A control that noticed only one of
        # them would be satisfied by rewording rather than by the claim being
        # made, and the wording has legitimately changed twice already.
        for phrase in ("and this is now the whole list", "and no further",
                       "only those", "limited to"):
            head = head.replace(phrase, "")
        return head
    damaged = _damage_preamble(rewrite)
    message = _preamble_control_refuses(damaged)
    assert "does not bound the runtime-proven claims" in message, message


def test_123b_the_closure_drops_the_boundary_it_did_not_cross() -> None:
    """While P6-ORA was open the preamble had to say its owner was unresolved.
    Closed, the equivalent honesty is the boundary: an all-green run proves the
    scenarios that ran, not the arms that cannot be reached, and a closure that
    stopped saying so would be over-claiming exactly where it costs most."""
    def rewrite(head):
        return head.replace("source-only", "settled").replace("static-only", "settled")
    damaged = _damage_preamble(rewrite)
    message = _preamble_control_refuses(damaged)
    assert "static-only boundary visible" in message, message


# ===========================================================================
# I. The Run-4 P6-ART file-lock defect
# ===========================================================================
def test_133_the_scenario_hashes_the_open_workbook_again() -> None:
    """The submitted shape, and Run 4 died on it: Get-FileHash cannot read a
    file the functional Excel instance is holding open."""
    damaged = _swap(_PHASE6,
                    "        $captured = @($ArtefactIdentity)\n",
                    "        $captured = @($ArtefactIdentity)\n"
                    "        $null = (Get-FileHash -LiteralPath "
                    "(Join-Path $TempRoot ([string]$Manifest.stage_b_filename)) "
                    "-Algorithm SHA256).Hash\n")
    _control("test_67", phase6=damaged)


def test_134_the_capture_is_weakened_to_the_build_directory() -> None:
    """Hashing build\\PCCM_stageB.xlsm would identify the directory the run was
    seeded from, not the disposable copy it consumed."""
    damaged = _swap(_PHASE6,
                    "        $path = Join-Path $TempRoot $item.Name\n        $hash = ''\n",
                    "        $path = Join-Path $BuildDir $item.Name\n        $hash = ''\n")
    _control("test_67", phase6=damaged)


def test_135_the_capture_can_throw() -> None:
    """It runs before Excel has started; an evidence step that aborts the run is
    a worse defect than the one it replaces."""
    damaged = _swap(_PHASE6, """        try {
            if (Test-Path -LiteralPath $path) {
                $hash = [string](Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
            } else {
                $problem = 'not found'
            }
        } catch {
            $problem = [string]$_.Exception.Message
        }
""",
                    """        $hash = [string](Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
""")
    _control("test_67", phase6=damaged)


def test_136_a_missing_capture_stops_failing_the_scenario() -> None:
    """No captured identity is not the same as an identity."""
    damaged = _swap(_PHASE6, "            ($captured.Count -eq 2) ('captured records: '",
                    "            ($captured.Count -ge 0) ('captured records: '")
    _control("test_67", phase6=damaged)


def test_137_a_blank_hash_is_reported_as_an_identity() -> None:
    """A record that carries no hash is a record of a failure, not evidence."""
    damaged = _swap(_PHASE6, "            $ok = ($item.Hash -match '^[0-9A-Fa-f]{64}$')",
                    "            $ok = $true")
    _control("test_67", phase6=damaged)  # noqa: E501 - the capture control, not the corpus one


def test_138_the_capture_is_not_bound_to_the_workbook_that_was_opened() -> None:
    """A hash taken before the open is evidence about THIS run only if the file
    Excel opened is the file that was hashed."""
    damaged = _swap(_PHASE6,
                    "             ([string]$executed[0].Path -ceq $openPath)) `",
                    "             ($true)) `")
    _control("test_67", phase6=damaged)


# ===========================================================================
# J. The Run-4 P6-FP3 preservation-set defect
# ===========================================================================
def test_139_the_hand_written_skip_list_comes_back() -> None:
    """The submitted shape. It named two of the seven rows WriteAttemptBlock
    rewrites, so Run 4 failed on status_evaluated_at being re-stamped."""
    damaged = _swap(_PHASE6, """        $durable = @($before['shared'].Keys | Where-Object {
            $rewritten -notcontains [string]$groups.$_ })
""",
                    """        $durable = @($before['shared'].Keys | Where-Object {
            ($_ -ne 'last_attempt_result') -and ($_ -ne 'last_attempt_detail') })
""")
    _control("test_68", phase6=damaged)


def test_140_the_projection_partition_is_traded_for_a_literal_list() -> None:
    """A list written here keeps asserting yesterday's contract."""
    damaged = _swap(_PHASE6, "        $groups = $SimInspection.sim_data.run_identity.groups\n",
                    "        $groups = @{}\n")
    _control("test_68", phase6=damaged)


def test_141_an_empty_durable_set_passes() -> None:
    """Preserving nothing is not preserving the durable rows."""
    damaged = _swap(_PHASE6,
                    "            ($durable.Count -gt 0) ('durable shared rows: '",
                    "            ($durable.Count -ge 0) ('durable shared rows: '")
    _control("test_68", phase6=damaged)


def test_142_the_restore_claim_is_no_longer_checked() -> None:
    """Without it the SameCell repair could regress and P6-FP3 stay green."""
    damaged = _swap(_PHASE6,
                    "            ($detail -notlike '*could not be restored*') $detail",
                    "            ($true) $detail")
    _control("test_68", phase6=damaged)


def test_143_the_derived_status_is_allowed_to_drift() -> None:
    """Nothing was published, so the derived state must not move."""
    damaged = _swap(_PHASE6, """        $null = Add-Check $list 'the derived simulation status is unchanged by the failed attempt' `
            (Test-SimSameValue -A $before['shared']['simulation_status'] `
                               -B $after['shared']['simulation_status']) `
            ('was ' + (Format-SimValue $before['shared']['simulation_status']) + ', now ' +
             (Format-SimValue $after['shared']['simulation_status']))
""", "")
    _control("test_68", phase6=damaged)


def test_144_the_candidate_claim_goes_back_to_being_a_note() -> None:
    """A note is not a check, and Run 4 recorded exactly that."""
    damaged = _swap(_PHASE6,
                    "            ((Get-Phase6ActiveBank -State $after) -cne $candidate) `",
                    "            ($true) `")
    _control("test_68", phase6=damaged)


# ===========================================================================
# K. The Step-0 evidence policy Step 13 overrode, and the D2 portability defect
# ===========================================================================
def test_145_exact_cross_language_digest_equality_is_restored() -> None:
    """The submitted rule, and the one Run 4 failed on. Step 0 §10.4 keeps the
    digest exact for SAME-RUNTIME replay; it never promised it across two
    languages, and it resolves one ULP in one iteration out of a thousand."""
    damaged = _swap(
        _PHASE6,
        "    Add-Note ($Label + ': result_digest oracle ' +\n",
        "    $null = Add-Check $List ($Label + ': published result_digest equals the oracle') `\n"
        "        (Test-SimExactText -Actual $publishedDigest "
        "-Expected ([string]$Measured.result_digest)) ''\n"
        "    Add-Note ($Label + ': result_digest oracle ' +\n")
    _control("test_69", phase6=damaged)


def test_146_a_floating_summary_row_returns_to_exact_equality() -> None:
    """§10.3 settled summary statistics at a relative bound with a scale-aware
    floor. Exact equality is a stronger rule than the accepted policy."""
    damaged = _swap(
        _PHASE6,
        "                (Test-Phase6WithinPolicy -Actual ([double]$actual) `\n"
        "                    -Expected ([double]$ladder.$rowKey) -Rule $rule -Scale $scale) `\n",
        "                (Test-SimExactDouble -Actual $actual "
        "-Expected ([double]$ladder.$rowKey)) `\n")
    _control("test_69", phase6=damaged)


def test_147_the_harness_spells_its_own_tolerance() -> None:
    """§10.1 gave the tolerance a single owner and it is not the harness. A
    literal here is a second authority, and it would drift from the first."""
    damaged = _swap(
        _PHASE6,
        "    if (($magnitude -gt 0) -and ($gap -le ([double]$Rule.relative * $magnitude))) { return $true }\n",
        "    if (($magnitude -gt 0) -and ($gap -le (1e-9 * $magnitude))) { return $true }\n")
    _control("test_10", phase6=damaged)


def test_148_the_scale_aware_floor_is_dropped() -> None:
    """Cancellation can drive a total near zero while every contribution that
    made it was large; a purely relative test is then unusable."""
    damaged = _swap(
        _PHASE6,
        "    if ([double]$Rule.absolute_floor -le 0) { return $false }\n"
        "    return ($gap -le ([double]$Rule.absolute_floor * [Math]::Abs($Scale)))\n",
        "    return $false\n")
    _control("test_10", phase6=damaged)


def test_149_the_same_runtime_digest_equality_is_weakened() -> None:
    """The one place §10.4 DOES promise digest equality, and it is not for
    weakening: two runs of one request inside one runtime."""
    damaged = _swap(
        _PHASE6,
        "        $null = Add-Check $list 'both runs published the same result digest' `\n"
        "            (Test-SimSameValue -A $digests[0] -B $digests[1]) `\n",
        "        $null = Add-Check $list 'both runs published the same result digest' `\n"
        "            ($true) `\n")
    _control("test_70", phase6=damaged)


def test_150_the_oracle_digest_dependency_returns_to_repeatability() -> None:
    """Run 4 proved P6-DET's own property and was marked red for this clause."""
    damaged = _swap(
        _PHASE6,
        "        # AND NOTHING CROSS-LANGUAGE.",
        "        $null = Add-Check $list 'and it is the digest the oracle predicted' `\n"
        "            (Test-SimExactText -Actual $digests[1] "
        "-Expected ([string]$OracleEvidence.measurements[0].result_digest)) ''\n"
        "        # AND NOTHING CROSS-LANGUAGE.")
    _control("test_70", phase6=damaged)


def test_151_the_host_sensitive_numbers_return_to_the_portable_authority() -> None:
    """D2: the same source produced 8C0D021F on Windows and 8C17DF7C on Linux,
    because the file carried Beta-PERT output and Cheng reaches libm."""
    def edit(document):
        document["parity_cases"][0]["expected_exact"]["result_digest"] = "0000000000000000"
    _control("test_48", cases=_json_mutation(_CASES, edit))


def test_152_the_authority_claims_a_portability_it_does_not_have() -> None:
    """Claiming invariance is not having it; the claim has to match the split."""
    def edit(document):
        document["portability"]["cross_platform_invariant"] = True
        document["measurements"][0]["summary"]["nominal"]["mean"] = 1.0
    _control("test_71", oracle=_json_mutation(_ORACLE, edit))


def test_153_the_evidence_is_generated_for_a_different_authority() -> None:
    """A pair assembled from two builds is not a pair, and the preflight has to
    refuse it before Excel rather than compare across it."""
    def edit(document):
        document["generated_for"]["sha256"] = "0" * 64
    _control("test_48", oracle=_json_mutation(_ORACLE, edit))


def test_154_the_preflight_stops_binding_the_two_artefacts() -> None:
    """Without the hash comparison the pair could come from two builds."""
    damaged = _swap(
        _PHASE6,
        "        (($authoritySha.Length -eq 64) -and\n"
        "         ($actualSha -ne '') -and ($actualSha -ieq $authoritySha)) `\n",
        "        ($true) `\n")
    _control("test_71", phase6=damaged)


def test_155_the_emitted_policy_drifts_from_the_step_0_record() -> None:
    """The constants in code are a COPY of §10.3, and a copy that stopped
    matching would be a second authority with a friendlier number."""
    def edit(document):
        document["comparison_policy"]["tolerances"]["summary_statistic"]["relative"] = 1e-6
    _control("test_31", cases=_json_mutation(_CASES, edit))


def test_156_the_policy_promises_a_cross_language_digest_again() -> None:
    """The qualifier Step 13 read past, restored as data this time."""
    def edit(document):
        document["comparison_policy"]["cross_language_digest_is_exact"] = True
    _control("test_31", cases=_json_mutation(_CASES, edit))


# ===========================================================================
# L. The superseded wording, and the host-local oracle's run binding
# ===========================================================================
_DOC = conformance.PCCM_ROOT / "docs" / "phase6_step13_gate_b.md"


def _wording_refuses(doc: str | None = None, phase6: str | None = None) -> str:
    """Run the current-wording detector over a damaged document or banner."""
    saved = (conformance.PCCM_ROOT, conformance.PHASE6)
    with tempfile.TemporaryDirectory(prefix="pccm-step13-wording-") as name:
        root = Path(name)
        (root / "docs").mkdir()
        (root / "docs" / "phase6_step13_gate_b.md").write_text(
            doc if doc is not None else _DOC.read_text(encoding="utf-8"),
            encoding="utf-8")
        conformance.PCCM_ROOT = root
        if phase6 is not None:
            target = root / conformance.PHASE6.name
            target.write_text(phase6, encoding="utf-8")
            conformance.PHASE6 = target
        try:
            conformance.test_45_the_current_wording_states_the_accepted_comparison_and_the_split()
        except AssertionError as error:
            return str(error)
        finally:
            conformance.PCCM_ROOT, conformance.PHASE6 = saved
    raise AssertionError("the mutation survived the current-wording control")


def _doc_swap(old: str, new: str) -> str:
    text = _DOC.read_text(encoding="utf-8")
    assert text.count(old) == 1, old[:80]
    return text.replace(old, new)


def test_157_the_exact_only_policy_returns_to_the_architecture() -> None:
    """The rule Step 13 wrote over the top of Step 0 §10."""
    damaged = _doc_swap(
        "#### Iterations, seed and the two comparison classes",
        "#### Iterations, seed and exactness\n\nThe comparison policy is **EXACT** "
        "and admits no tolerance.\n")
    message = _wording_refuses(doc=damaged)
    assert "admits" in message or "EXACT" in message, message


def test_158_the_matrix_claims_excel_equals_the_oracle_again() -> None:
    """Digest equality across two languages, back in the scenario definition."""
    line = [l for l in _DOC.read_text(encoding="utf-8").splitlines()
            if l.startswith("| `P6-ORA` |")][0]
    damaged = _doc_swap(
        line,
        "| `P6-ORA` | **Excel equals the oracle** — digest, seeds, versions and "
        "the full ladder | cells vs corpus |")
    message = _wording_refuses(doc=damaged)
    assert "equals the oracle" in message, message


def test_159_the_matrix_restores_the_oracle_digest_to_repeatability() -> None:
    """P6-DET proves a same-runtime replay property and nothing else."""
    line = [l for l in _DOC.read_text(encoding="utf-8").splitlines()
            if l.startswith("| `P6-DET` |")][0]
    damaged = _doc_swap(
        line,
        "| `P6-DET` | the same inputs and seed twice produce the same digest, "
        "and it is the oracle's | cells |")
    message = _wording_refuses(doc=damaged)
    assert "oracle" in message, message


def test_160_the_source_banner_denies_every_execution_again() -> None:
    """Runs 1-4 have run, and Run 4 executed the behavioural matrix."""
    damaged = _swap(
        _PHASE6,
        "    WHAT HAS EXECUTED. Runs 1-6 have run.",
        "    NOTHING HERE HAS BEEN EXECUTED. As submitted, no Windows run has been\n"
        "    made. Runs 1-6 have run.")
    message = _wording_refuses(phase6=damaged)
    assert "NOTHING HERE HAS BEEN EXECUTED" in message, message


def test_161_the_banner_sends_every_expected_value_to_one_artefact() -> None:
    """True before D2, false after it: the floating ladder and the diagnostic
    digest come from the host-local companion."""
    damaged = _swap(
        _PHASE6,
        "    * it restates no expected simulation number, and there are TWO artefacts it\n",
        "    * it restates no expected simulation number. Every digest, seed, ladder\n"
        "      value and bound comes from build/phase6_gate_b_cases.json.\n"
        "    * and there are TWO artefacts it\n")
    damaged = damaged.replace("phase6_gate_b_oracle_local.json,\n"
                              "      the HOST-LOCAL oracle measurements", "the same file")
    message = _wording_refuses(phase6=damaged)
    assert "one artefact" in message or "host-local companion" in message, message


def test_162_the_pre_excel_gate_is_misnamed_as_p6_pre() -> None:
    """P6-PRE executes later, inside the live session; the artefact gate is
    PRE6, and confusing them misdescribes where a refusal happens."""
    damaged = _doc_swap(
        "**Binding.** `PRE6` — `Invoke-Phase6CoveragePreflight`, the pure artefact gate\n"
        "that runs before Excel is started, and NOT `P6-PRE`, which executes later inside\n"
        "the live session — refuses if the file is missing",
        "**Binding.** `P6-PRE` refuses, pre-Excel, if the file is missing")
    message = _wording_refuses(doc=damaged)
    assert "P6-PRE" in message, message


# ---- the run binding -------------------------------------------------------
def test_163_the_oracle_carries_no_real_commit_identity() -> None:
    """`unavailable` is what the builder writes when git cannot be read. It is
    honest, and it is not a Gate-B evidence package."""
    def edit(document):
        document["source_revision"] = "unavailable"
    _control("test_72", oracle=_json_mutation(_ORACLE, edit))


def test_164_the_oracle_names_a_commit_this_repository_does_not_have() -> None:
    """A well-formed SHA that nothing can be attributed to."""
    def edit(document):
        document["source_revision"] = "0" * 40
    _control("test_72", oracle=_json_mutation(_ORACLE, edit))


def test_165_the_run_binding_is_weakened_to_a_prefix() -> None:
    """Two representations of one commit make the binding a string-formatting
    question. This is the shape a stale short-SHA oracle would slip through."""
    damaged = _swap(
        _PHASE6,
        "         ($oracleRevision -ceq $HarnessCommit)) `\n",
        "         ($oracleRevision -like ($HarnessCommit.Substring(0, 7) + '*'))) `\n")
    _control("test_72", phase6=damaged)


def test_166_the_pre_excel_gate_stops_comparing_the_revision() -> None:
    """Without it a stale oracle_local from an earlier harness commit passes:
    a harness-only commit does not move the portable authority, so its SHA-256
    still matches and every other check is satisfied."""
    damaged = _swap(
        _PHASE6,
        "    $null = Add-Check $list 'the oracle evidence was generated at THIS run''s harness commit' `\n"
        "        (($oracleRevision -match '^[0-9a-f]{40}$') -and\n"
        "         ($HarnessCommit -match '^[0-9a-f]{40}$') -and\n"
        "         ($oracleRevision -ceq $HarnessCommit)) `\n",
        "    $null = Add-Check $list 'the oracle evidence was generated at THIS run''s harness commit' `\n"
        "        ($true) `\n")
    _control("test_72", phase6=damaged)


def test_167_the_runtime_scenario_only_prints_the_revision() -> None:
    """P6-ART recording a revision is attribution; re-asserting it is a binding."""
    damaged = _swap(
        _PHASE6,
        "            (([string]$OracleEvidence.source_revision -match '^[0-9a-f]{40}$') -and\n"
        "             ([string]$OracleEvidence.source_revision -ceq [string]$HarnessCommit)) `\n",
        "            ($true) `\n")
    _control("test_72", phase6=damaged)


def test_168_the_pair_disagrees_on_the_supplied_seed() -> None:
    def edit(document):
        document["supplied_seed"] = 99999
    _control("test_72", oracle=_json_mutation(_ORACLE, edit))


def test_169_the_pair_disagrees_on_the_schema_version() -> None:
    def edit(document):
        document["schema_version"] = 99
    _control("test_72", oracle=_json_mutation(_ORACLE, edit))


def test_170_the_oracle_names_a_different_case_authority() -> None:
    def edit(document):
        document["generated_for"]["authority"] = "phase5_cases.json"
    _control("test_72", oracle=_json_mutation(_ORACLE, edit))


def test_171_the_oracle_names_a_different_policy_authority() -> None:
    """The measurements and the rule they are compared under must come from one
    settled policy, not two."""
    def edit(document):
        document["evidence_policy_authority"] = "docs/phase6_plan.md §15.1"
    _control("test_72", oracle=_json_mutation(_ORACLE, edit))


# ===========================================================================
# M. A commit id is not a set of bytes
# ===========================================================================
def test_172_the_generation_records_no_clean_tree_fact_at_all() -> None:
    """A revision alone cannot see the generate-dirty-then-revert path."""
    def edit(document):
        del document["source_tree_clean"]
    _control("test_73", oracle=_json_mutation(_ORACLE, edit))


def test_173_the_oracle_was_generated_from_a_dirty_tracked_tree() -> None:
    """The whole point of recording it at generation: a tracked builder source
    was modified, the measurements were produced from it, and reverting the file
    afterwards cannot retract what the artefact already says."""
    def edit(document):
        document["source_tree_clean"] = False
    _control("test_73", oracle=_json_mutation(_ORACLE, edit))


def test_174_the_generation_hardcodes_a_clean_tree() -> None:
    """A verdict that was never established is not a verdict."""
    emit = conformance.PCCM_ROOT / "builder" / "pccm_builder" / "sim_emit.py"
    original = emit.read_text(encoding="utf-8")
    damaged = original.replace("    return revision, diff.returncode == 0",
                               "    return revision, True", 1)
    assert damaged != original
    saved = conformance.PCCM_ROOT
    with tempfile.TemporaryDirectory(prefix="pccm-step13-emit-") as name:
        root = Path(name)
        (root / "builder" / "pccm_builder").mkdir(parents=True)
        (root / "builder" / "pccm_builder" / "sim_emit.py").write_text(
            damaged, encoding="utf-8")
        for relative in ("build/phase6_gate_b_oracle_local.json",):
            (root / relative).parent.mkdir(parents=True, exist_ok=True)
            (root / relative).write_text(_ORACLE, encoding="utf-8")
        conformance.PCCM_ROOT = root
        try:
            refused = False
            try:
                conformance.test_73_the_evidence_is_bound_to_head_bytes_and_not_only_to_a_commit_id()
            except AssertionError as error:
                refused = True
                assert "without having established" in str(error), error
        finally:
            conformance.PCCM_ROOT = saved
    assert refused, "the builder may report a clean tree it never established"


def test_175_the_generation_stops_comparing_against_head() -> None:
    """`rev-parse HEAD` alone is the defect this settles."""
    emit = conformance.PCCM_ROOT / "builder" / "pccm_builder" / "sim_emit.py"
    original = emit.read_text(encoding="utf-8")
    damaged = original.replace('"diff", "--quiet", "HEAD"', '"rev-parse", "HEAD"', 1)
    assert damaged != original
    saved = conformance.PCCM_ROOT
    with tempfile.TemporaryDirectory(prefix="pccm-step13-emit2-") as name:
        root = Path(name)
        (root / "builder" / "pccm_builder").mkdir(parents=True)
        (root / "builder" / "pccm_builder" / "sim_emit.py").write_text(
            damaged, encoding="utf-8")
        (root / "build").mkdir()
        (root / "build" / "phase6_gate_b_oracle_local.json").write_text(
            _ORACLE, encoding="utf-8")
        conformance.PCCM_ROOT = root
        try:
            refused = False
            try:
                conformance.test_73_the_evidence_is_bound_to_head_bytes_and_not_only_to_a_commit_id()
            except AssertionError as error:
                refused = True
                assert "against HEAD" in str(error), error
        finally:
            conformance.PCCM_ROOT = saved
    assert refused, "the builder no longer compares the tracked tree against HEAD"


def test_176_the_pre_excel_gate_stops_checking_the_running_tree() -> None:
    """A tracked harness file modified after Stage A leaves HEAD unchanged, so
    every revision comparison still passes while other bytes execute."""
    damaged = _swap(
        _PHASE6,
        "    $null = Add-Check $list 'the tracked pccm tree being executed matches the harness commit' `\n"
        "        $treeClean $treeDetail\n",
        "    $null = Add-Check $list 'the tracked pccm tree being executed matches the harness commit' `\n"
        "        $true $treeDetail\n")
    _control("test_73", phase6=damaged)


def test_177_the_runtime_scenario_stops_re_asserting_the_running_tree() -> None:
    """P6-ART reports the harness commit; the evidence has to stay attributable
    to it inside the session, not only at the gate."""
    damaged = _swap(
        _PHASE6,
        "        $null = Add-Check $list 'the tracked pccm tree that ran matches the harness commit' `\n"
        "            $runtimeClean $runtimeDetail\n",
        "        $lines += ('runtime tracked tree clean: ' + [string]$runtimeClean)\n")
    _control("test_73", phase6=damaged)


def test_178_the_pre_excel_gate_stops_requiring_a_clean_generation() -> None:
    """Generation-time and runtime are different questions, and dropping the
    first restores the generate-dirty-then-revert path in full."""
    damaged = _swap(
        _PHASE6,
        "    $null = Add-Check $list 'the oracle evidence was generated from a clean tracked tree' `\n"
        "        ([bool]$oracle.source_tree_clean) `\n",
        "    $null = Add-Check $list 'the oracle evidence was generated from a clean tracked tree' `\n"
        "        ($true) `\n")
    _control("test_73", phase6=damaged)


def test_179_the_running_tree_check_narrows_to_a_subpath() -> None:
    """`pccm/src` would leave the builder, the policy and the whole Step-13
    PowerShell outside the statement - every source capable of changing the run."""
    damaged = _PHASE6.replace("diff --quiet HEAD -- 'pccm' 2>$null",
                              "diff --quiet HEAD -- 'pccm/src' 2>$null")
    assert damaged != _PHASE6
    _control("test_73", phase6=damaged)


# ===========================================================================
# N. PRE6 is not P6-PRE, and the driver banner is current documentation
# ===========================================================================
def test_180_the_preflight_exception_manufactures_a_p6_pre_result() -> None:
    """The submitted shape. An exception before Excel recorded a verdict for a
    live-session scenario that had not been reached, on the one path where
    nothing else could contradict it."""
    damaged = _swap(
        _HARNESS,
        "    Write-Host ('PRE6 (the Phase-6 artefact preflight) raised: ' + (Format-Err $_)) -ForegroundColor Red\n",
        "    Add-Phase6Result 'P6-PRE' 'Phase-6 artefact preflight' 'FAIL' (Format-Err $_)\n")
    _control("test_74", harness=damaged)


def test_181_any_phase6_result_is_recorded_before_excel_exists() -> None:
    """Not only P6-PRE: no Phase-6 scenario verdict can be honest before the
    session that produces the evidence has started."""
    damaged = _swap(
        _HARNESS,
        "    Write-Host ('PRE6 (the Phase-6 artefact preflight) raised: ' + (Format-Err $_)) -ForegroundColor Red\n",
        "    Add-Phase6Result 'P6-ART' 'Phase-6 artefact preflight' 'FAIL' (Format-Err $_)\n")
    _control("test_74", harness=damaged)


def test_182_the_preflight_exception_stops_aborting() -> None:
    """A preflight that raises and continues would start Excel on artefacts it
    could not read."""
    damaged = _swap(
        _HARNESS,
        "    Write-Host ('PRE6 (the Phase-6 artefact preflight) raised: ' + (Format-Err $_)) -ForegroundColor Red\n"
        "    exit 1\n"
        "}\n",
        "    Write-Host ('PRE6 (the Phase-6 artefact preflight) raised: ' + (Format-Err $_)) -ForegroundColor Red\n"
        "}\n")
    _control("test_74", harness=damaged)


def test_183_pre6_is_promoted_to_a_phase6_scenario_id() -> None:
    """PRE6 is a preflight label. Adding it to the scenario vocabulary would put
    a pre-Excel gate inside the set P6-FIN proves complete."""
    damaged = _swap(
        _PHASE6,
        "        'P6-SU', 'P6-XX', 'P6-LDG', 'P6-FIN'\n",
        "        'P6-SU', 'P6-XX', 'P6-LDG', 'P6-FIN', 'PRE6'\n")
    _control("test_74", phase6=damaged)


def test_184_the_banner_claims_p5_p4_checks_the_final_matrix() -> None:
    """Y and Z are POST-SESSION, so what P5-P4 can check is the derived
    prerequisite partition; the final 35/35 is P5-FIN's."""
    damaged = _swap(
        _HARNESS,
        "    matrix above is unchanged and remains mandatory, but it is not complete when\n",
        "    matrix above is unchanged and remains mandatory: P5-P4 checks it reached\n"
        "    35/35 with 0 FAIL and 0 SKIP.\n"
        "    It is not complete when\n")
    _control("test_75", harness=damaged)


def test_185_the_banner_denies_every_phase5_gate_b_run() -> None:
    """Phase 5 Gate B has executed on real Excel over several Windows runs, and
    Phase 5 was accepted on that evidence."""
    damaged = _swap(
        _HARNESS,
        "    PHASE 5 GATE B IS CLOSED.",
        "    NO PHASE-5 GATE-B RUN HAS BEEN MADE. This harness extension is source under\n"
        "    independent review; no Excel COM session has been started for it.\n"
        "    PHASE 5 GATE B IS CLOSED.")
    _control("test_75", harness=damaged)


def test_188_the_pre6_comment_describes_two_authorities_again() -> None:
    """The pre-D2 world: one address projection and one corpus. The gate reads
    three artefacts now, and the third is host-local by construction."""
    damaged = _swap(
        _HARNESS,
        "# THREE ARTEFACTS, AND THEY ARE NOT INTERCHANGEABLE. The Step-13 scenarios read\n",
        "# The inspection projection and the parity corpus are the only two authorities\n"
        "# the Step-13 scenarios read. Also,\n")
    _control("test_45", harness=damaged)


def test_189_the_source_control_docstring_denies_run_4() -> None:
    """Run 4 executed the Phase-6 behavioural matrix. What is unexercised is
    everything settled after it, which is a different statement."""
    conformance_path = Path(conformance.__file__)
    original = conformance_path.read_text(encoding="utf-8")
    damaged = original.replace(
        "WHAT A STATIC CONTROL CAN AND CANNOT SETTLE, which is a property and not a date.",
        "What it does NOT establish is anything about behaviour. Whether Excel agrees is\n"
        "Gate B's, on Windows, and that run has not been made.\n", 1)
    assert damaged != original
    # THE DOCSTRING IS READ FROM THE FILE, so the mutation is installed as a
    # damaged copy of this module's own source and the control re-run over it.
    with tempfile.TemporaryDirectory(prefix="pccm-step13-docstring-") as name:
        target = Path(name) / conformance_path.name
        target.write_text(damaged, encoding="utf-8")
        saved = conformance.__file__
        conformance.__file__ = str(target)
        try:
            refused = False
            try:
                conformance.test_45_the_current_wording_states_the_accepted_comparison_and_the_split()
            except AssertionError as error:
                refused = True
                assert "denies the run" in str(error), error
        finally:
            conformance.__file__ = saved
    assert refused, "the module docstring may deny Run 4's behavioural evidence"


# ===========================================================================
# O. Invariant artefacts are BYTES, not platform-translated text
# ===========================================================================
def _builder_refuses(relative: str, damaged: str, fragment: str) -> None:
    """Re-run the serialisation control over a damaged builder module."""
    saved = conformance.PCCM_ROOT
    with tempfile.TemporaryDirectory(prefix="pccm-serialisation-") as name:
        root = Path(name)
        (root / "builder" / "pccm_builder").mkdir(parents=True)
        for module in ("sim_emit.py", "sim_inspection.py"):
            source = (saved / "builder" / "pccm_builder" / module)
            (root / "builder" / "pccm_builder" / module).write_text(
                damaged if module == relative else source.read_text(encoding="utf-8"),
                encoding="utf-8")
        conformance.PCCM_ROOT = root
        try:
            refused = False
            try:
                conformance.test_76_the_invariant_artefacts_are_written_as_bytes_not_translated_text()
            except AssertionError as error:
                refused = True
                assert fragment in str(error), error
        finally:
            conformance.PCCM_ROOT = saved
    assert refused, f"the mutation survived: {relative} may translate newlines"


def _module(name: str) -> str:
    return (conformance.PCCM_ROOT / "builder" / "pccm_builder" / name).read_text(
        encoding="utf-8")


def test_190_the_inspection_projection_returns_to_text_mode() -> None:
    """The submitted shape, and Windows produced CRLF from it."""
    original = _module("sim_inspection.py")
    damaged = original.replace(
        '    write_lf_artifact(path, json.dumps(document, indent=2, sort_keys=False) + "\\n")',
        '    path.write_text(json.dumps(document, indent=2, sort_keys=False) + "\\n",\n'
        '                    encoding="utf-8")', 1)
    assert damaged != original
    _builder_refuses("sim_inspection.py", damaged, "text mode")


def test_191_the_portable_authority_returns_to_text_mode() -> None:
    original = _module("sim_emit.py")
    damaged = original.replace(
        "    cases_bytes = write_lf_artifact(\n"
        '        cases_path, json.dumps(portable, indent=2, sort_keys=False) + "\\n")',
        '    cases_text = json.dumps(portable, indent=2, sort_keys=False) + "\\n"\n'
        '    cases_path.write_text(cases_text, encoding="utf-8")\n'
        "    cases_bytes = cases_text.encode(\"utf-8\")", 1)
    assert damaged != original
    _builder_refuses("sim_emit.py", damaged, "text mode")


def test_192_the_phase6_cases_corpus_returns_to_text_mode() -> None:
    """It carries a pinned SHA-256 too, so its bytes are a claim as well."""
    original = _module("sim_emit.py")
    damaged = original.replace(
        "    write_lf_artifact(cases_path, render_sim_cases_json(spec, sim, inputs, calc))",
        '    cases_path.write_text(render_sim_cases_json(spec, sim, inputs, calc),\n'
        '                          encoding="utf-8")', 1)
    assert damaged != original
    _builder_refuses("sim_emit.py", damaged, "text mode")


def test_193_the_authority_hash_stops_describing_the_emitted_bytes() -> None:
    """Hashing the string the emitter HAD is what made `generated_for.sha256`
    name a file that never existed on Windows."""
    def edit(document):
        document["generated_for"]["sha256"] = "f" * 64
    _control("test_76", oracle=_json_mutation(_ORACLE, edit))


def test_194_a_carriage_return_reaches_the_inspection_projection() -> None:
    damaged = _INSPECTION.replace("\n", "\r\n", 1)
    assert damaged != _INSPECTION
    _control("test_76", inspection=damaged)


def test_195_a_carriage_return_reaches_the_portable_authority() -> None:
    damaged = _CASES.replace("\n", "\r\n", 1)
    assert damaged != _CASES
    _control("test_76", cases=damaged)


def test_196_a_bom_is_emitted() -> None:
    """A BOM is three bytes that move every pinned hash and nothing else."""
    _control("test_76", cases="\ufeff" + _CASES)


def test_197_the_final_newline_is_dropped() -> None:
    _control("test_76", cases=_CASES.rstrip("\n"))


# ===========================================================================
# P. Two classes of production module, and two kinds of identity
# ===========================================================================
def test_198_the_generated_projection_returns_to_the_tracked_blob_loop() -> None:
    """Run 5's defect, exactly. modSimContract has no path in src/vba, so the
    comparison passes a blank against a blank on every commit."""
    damaged = _swap(
        _PHASE6,
        "    return @('modSimRng', 'modSimSample', 'modSimEngine',\n",
        "    return @('modSimContract', 'modSimRng', 'modSimSample', 'modSimEngine',\n")
    _control("test_52", phase6=damaged)


def test_199_the_generated_projection_is_dropped_from_identity_checking() -> None:
    """Eight production modules are exercised; seven proved is not all of them."""
    damaged = _swap(
        _PHASE6,
        "            $null = Add-Check $list `\n"
        "                ('the ' + $generatedName + ' projection is the one baseline ' + $baseline + ' produces') `\n"
        "                ($observed -ceq $accepted) `\n",
        "            $null = Add-Check $list `\n"
        "                ('the ' + $generatedName + ' projection is the one baseline ' + $baseline + ' produces') `\n"
        "                ($true) `\n")
    _control("test_77", phase6=damaged)


def test_200_a_manifest_entry_that_is_not_generated_is_accepted() -> None:
    """`generated: true` is what sends Stage B to build/vba rather than src/vba;
    an entry without it describes a different import."""
    damaged = _swap(
        _PHASE6,
        "                (($entries.Count -eq 1) -and ([bool]$entries[0].generated)) `\n",
        "                (($entries.Count -eq 1)) `\n")
    _control("test_77", phase6=damaged)


def test_201_zero_or_duplicate_manifest_entries_are_accepted() -> None:
    damaged = _swap(
        _PHASE6,
        "                ($entries.Count -eq 1) ('entries: ' + $entries.Count)\n",
        "                ($entries.Count -ge 0) ('entries: ' + $entries.Count)\n")
    _control("test_77", phase6=damaged)


def test_202_the_generated_module_is_resolved_from_the_source_tree() -> None:
    """Resolving it from src/vba would hash a file Stage B never imported - and
    in this project that file does not exist at all."""
    damaged = _swap(
        _PHASE6,
        "            $generatedDir = Join-Path $TempRoot (Split-Path -Leaf ([string]$Manifest.vba.generated_dir))\n",
        "            $generatedDir = Join-Path $TempRoot 'src'\n")
    _control("test_77", phase6=damaged)


def test_203_the_expected_identity_becomes_head_derived() -> None:
    """The property that matters: a changed renderer must not be able to change
    the projection and then bless its own changed output."""
    original = conformance_source = Path(conformance.__file__).read_text(encoding="utf-8")
    damaged = original.replace(
        "    derived = _baseline_projection_identity()",
        "    from pccm_builder.artifact_io import canonical_module_identity as _c\n"
        "    derived = _c((BUILD / 'vba' / 'modSimContract.bas').read_bytes())", 1)
    assert damaged != original
    assert "_baseline_projection_identity()" not in damaged.split(
        "def test_77")[1].split("\ndef ")[0], (
        "the mutation did not remove the baseline derivation"
    )
    # THE SHAPE IS THE PROOF: a control that derives both sides from HEAD cannot
    # distinguish a moved renderer from an unmoved one, whatever it then asserts.
    assert "PRODUCTION_BASELINE" in original.split("def _baseline_projection_identity")[1].split("\ndef ")[0]


def test_204_the_raw_windows_hash_is_pinned_as_the_identity() -> None:
    """The .bas is text mode by design, so its raw bytes are host-dependent; the
    Windows diagnostic SHA is evidence about one host, not an identity."""
    damaged = _swap(
        _PHASE6,
        "    return 'daa4d27889c30eadb2ab892bcfa4e6f6bab8a137aae79a01a8d8f1e8e1c215ac'\n",
        "    return 'cc74eec48d3f1d9b5d66b4441cbb6540593bbb89329304a0405dff425a3403c2'\n")
    _control("test_77", phase6=damaged)


def test_205_a_hand_written_module_blob_check_is_weakened() -> None:
    """The seven are not touched by any of this."""
    damaged = _swap(
        _PHASE6,
        "                    ((-not [string]::IsNullOrWhiteSpace($accepted)) -and ($current -ceq $accepted)) `\n",
        "                    ($true) `\n")
    _control("test_52", phase6=damaged)


def test_206_the_canonicaliser_accepts_a_bom() -> None:
    """Stripping a BOM would hide a real difference in the projection."""
    from pccm_builder.artifact_io import (  # noqa: PLC0415
        ArtifactSerialisationError,
        canonical_module_identity,
    )
    try:
        canonical_module_identity(b"\xef\xbb\xbfAttribute VB_Name = \"m\"\n")
    except ArtifactSerialisationError:
        return
    raise AssertionError("the canonicaliser accepted a BOM")


def test_207_the_canonicaliser_accepts_a_missing_final_newline() -> None:
    from pccm_builder.artifact_io import (  # noqa: PLC0415
        ArtifactSerialisationError,
        canonical_module_identity,
    )
    try:
        canonical_module_identity(b"Attribute VB_Name = \"m\"")
    except ArtifactSerialisationError:
        return
    raise AssertionError("the canonicaliser accepted a module with no final newline")


def _canonicaliser_refuses(damaged: str, why: str) -> None:
    """Re-run the projection control against a damaged canonicaliser.

    The control calls the FUNCTION, not its source text - which is the whole
    point of the rewrite - so the damage is installed by compiling the modified
    module and swapping the function into place. Its refusal class is rebound to
    the real one so a refusal it does raise is still caught as a refusal, and a
    mutation cannot pass by raising something the control was not watching for.
    """
    from pccm_builder import artifact_io  # noqa: PLC0415

    namespace: dict[str, object] = {"__name__": "pccm_builder.artifact_io_damaged"}
    exec(compile(damaged, "artifact_io_damaged", "exec"), namespace)  # noqa: S102
    namespace["ArtifactSerialisationError"] = artifact_io.ArtifactSerialisationError
    saved = artifact_io.canonical_module_identity
    artifact_io.canonical_module_identity = namespace["canonical_module_identity"]
    try:
        refused = False
        try:
            conformance.test_77_the_generated_projection_has_a_baseline_bound_identity()
        except AssertionError as error:
            refused = True
            assert why in str(error), error
    finally:
        artifact_io.canonical_module_identity = saved
    assert refused, f"the canonicaliser accepts {why}"


def test_208_the_canonicaliser_normalises_more_than_line_endings() -> None:
    """Anything broader lets a real change pass as the same module."""
    original = (conformance.PCCM_ROOT / "builder" / "pccm_builder" / "artifact_io.py"
                ).read_text(encoding="utf-8")
    damaged = original.replace(
        '    canonical = data.replace(b"\\r\\n", b"\\n")\n',
        '    canonical = data.replace(b"\\r\\n", b"\\n").lower()\n', 1)
    assert damaged != original
    # A broader normalisation changes the identity of EVERY module, so the
    # baseline comparison is what refuses first. Either refusal is the same
    # finding; the fragment names the one that actually fires.
    _canonicaliser_refuses(damaged, "projection identity")



def test_209_a_bare_cr_is_canonicalised_to_lf_again() -> None:
    """The submitted defect. `.replace(b"\\r", b"\\n")` admitted a third line-ending
    representation, so a module whose every LF had become a lone CR hashed
    identically to the accepted projection."""
    original = (conformance.PCCM_ROOT / "builder" / "pccm_builder" / "artifact_io.py"
                ).read_text(encoding="utf-8")
    damaged = original.replace(
        '    stray = data.replace(b"\\r\\n", b"")\n'
        '    if b"\\r" in stray:\n',
        '    stray = data.replace(b"\\r\\n", b"")\n'
        '    if False:\n', 1).replace(
        '    canonical = data.replace(b"\\r\\n", b"\\n")\n',
        '    canonical = data.replace(b"\\r\\n", b"\\n").replace(b"\\r", b"\\n")\n', 1)
    assert damaged != original
    assert 'replace(b"\\r", b"\\n")' in damaged
    _canonicaliser_refuses(damaged, "a bare CR as a line ending")


def test_210_the_bare_cr_refusal_is_removed_outright() -> None:
    """Without the refusal the replace is not even needed: a lone CR simply
    survives into the hash, and a corrupted module gets its own identity rather
    than being refused."""
    original = (conformance.PCCM_ROOT / "builder" / "pccm_builder" / "artifact_io.py"
                ).read_text(encoding="utf-8")
    damaged = original.replace(
        '    stray = data.replace(b"\\r\\n", b"")\n'
        '    if b"\\r" in stray:\n',
        '    stray = data.replace(b"\\r\\n", b"")\n'
        '    if False:\n', 1)
    assert damaged != original
    _canonicaliser_refuses(damaged, "a bare CR inside a line")


def test_211_the_harness_canonicaliser_converts_a_lone_cr() -> None:
    """The PowerShell side has to hold the same rule; a harness that mapped a
    lone CR onto LF would accept on Windows what Python refuses here."""
    damaged = _swap(
        _PHASE6,
        "            throw ('the generated module carries a carriage return that is not part ' +\n"
        "                   'of a CRLF; LF and CRLF are the accepted representations and a ' +\n"
        "                   'bare CR is not a third one')\n",
        "            $null = $canonical.Add([byte]0x0A)\n")
    _control("test_77", phase6=damaged)


# ===========================================================================
# Q. Closure is a claim, and a claim needs a control
# ===========================================================================
def _closure_refuses(rewrite, fragment: str) -> None:
    """Run the closure and preamble controls over a damaged document."""
    original = _DOC.read_text(encoding="utf-8")
    damaged = rewrite(original)
    assert damaged != original, "the mutation changed nothing"
    saved = conformance.PCCM_ROOT
    with tempfile.TemporaryDirectory(prefix="pccm-closure-") as name:
        root = Path(name)
        (root / "docs").mkdir()
        (root / "docs" / "phase6_step13_gate_b.md").write_text(damaged, encoding="utf-8")
        conformance.PCCM_ROOT = root
        try:
            refused = []
            for control in (
                conformance.test_78_the_closure_states_what_run_6_established_and_no_more,
                conformance.test_66_the_active_preamble_states_what_has_run_and_what_has_not,
            ):
                try:
                    control()
                except AssertionError as error:
                    refused.append(str(error))
        finally:
            conformance.PCCM_ROOT = saved
    assert refused, "the mutation survived both closure controls"
    assert any(fragment in message for message in refused), (fragment, refused)


def test_212_the_closing_run_is_recorded_without_having_executed() -> None:
    """`<this commit>` is the placeholder for a run that has not happened."""
    _closure_refuses(lambda doc: doc.replace("| **Run 6** | `a3924e0` |",
                                             "| **Run 6** | `<this commit>` |", 1),
                     "recorded as all green against")


def test_213_the_closing_row_records_a_failure() -> None:
    _closure_refuses(lambda doc: doc.replace("103 passed, 0 failed, 0 skipped. Phase-4",
                                             "103 passed, 1 failed, 0 skipped. Phase-4", 1),
                     # The tally demand refuses first; both messages name the
                     # same finding, and this is the one that fires.
                     "103 passed")


def test_214_the_closing_row_records_a_skip() -> None:
    _closure_refuses(lambda doc: doc.replace("103 passed, 0 failed, 0 skipped. Phase-4",
                                             "103 passed, 0 failed, 2 skipped. Phase-4", 1),
                     "103 passed")


def test_215_the_phase6_tally_is_less_than_the_matrix() -> None:
    _closure_refuses(lambda doc: doc.replace("Phase-6 29/29. `P6-ART` PASS",
                                             "Phase-6 27/29. `P6-ART` PASS", 1),
                     "does not state Phase-6 29/29")


def test_216_a_settled_scenario_is_still_reported_open() -> None:
    """P6-ART, P6-ORA, P6-DET and P6-FP3 were each open at some point; a closure
    that still says so is describing a state the ledger says it left."""
    _closure_refuses(
        lambda doc: doc.replace(
            "`P6-ORA` on the\naccepted tolerance rule",
            "`P6-ORA` remains OPEN and unresolved; the\naccepted tolerance rule", 1),
        "as open")


def test_217_a_settled_scenario_loses_its_pass_record() -> None:
    _closure_refuses(
        lambda doc: doc.replace(
            "P6-LDG                             PASS  (28 scenario results, 0 duplicates)\n",
            "", 1),
        "is not recorded as passing")


def test_218_the_status_line_stops_stating_closure() -> None:
    _closure_refuses(lambda doc: doc.replace("**Status: CLOSED.", "**Status: OPEN.", 1),
                     "does not state closure")


def test_219_the_closure_does_not_name_the_baseline_it_closed_against() -> None:
    _closure_refuses(
        lambda doc: doc.replace(
            "The production baseline remains `79e4600`.**",
            "The production baseline is unchanged.**", 1),
        "does not name the production baseline")


def test_220_a_static_only_subject_is_dropped_by_the_all_green_run() -> None:
    """An all-green run proves the scenarios that ran, not the arms that cannot
    be reached. Every subject removed here was never induced."""
    # THE ROW INSIDE THE BOUNDARY, not the first row anywhere. `FinalCommit`
    # also names a row in the scenario matrix, and removing that one leaves the
    # boundary intact - a mutation that damaged the wrong table would look like
    # a surviving mutation and send the next reader after the wrong control.
    document = _DOC.read_text(encoding="utf-8")
    boundary = document.split("## 5.")[1].split("\n## ")[0]
    for subject, _ in conformance.STATIC_ONLY_SUBJECTS:
        rows = [line for line in boundary.splitlines()
                if line.startswith("|") and subject in line]
        assert len(rows) == 1, (subject, len(rows))
        _closure_refuses(lambda doc, row=rows[0]: doc.replace(row + "\n", "", 1),
                         "has been dropped from the static-only list")


def _occurrences(text: str, needle: str) -> list[int]:
    found, at = [], text.find(needle)
    while at != -1:
        found.append(at)
        at = text.find(needle, at + 1)
    return found


def test_220b_a_static_only_row_is_weakened_rather_than_dropped() -> None:
    """Deleting a row is the obvious move and the least likely one. A row that
    keeps its subject and loses its qualifier still LOOKS like the exclusion it
    used to be - `ClearPending` without "after a known CONSUMED" is a different,
    broader claim, and `SharedReadRaised` without `ReadRaised` silently narrows
    the pair to one of its two members."""
    document = _DOC.read_text(encoding="utf-8")
    boundary = document.split("## 5.")[1].split("\n## ")[0]
    for subject, qualifiers in conformance.STATIC_ONLY_SUBJECTS:
        row = [line for line in boundary.splitlines()
               if line.startswith("|") and subject in line][0]
        for qualifier in qualifiers:
            # THE OCCURRENCE THAT IS NOT INSIDE THE SUBJECT. `ReadRaised` is a
            # substring of `SharedReadRaised`, and the qualifier can sit either
            # side of the subject in the row, so the occurrence to remove is
            # chosen by span: deleting the subject itself would exercise the
            # dropped-row arm instead of the narrowing one, which is exactly the
            # case this mutation exists for.
            cells = row.split("|")
            cell = cells[1]
            subject_spans = [(m, m + len(subject))
                             for m in _occurrences(cell, subject)]
            damaged_row = None
            for start in _occurrences(cell, qualifier):
                end = start + len(qualifier)
                if any(start >= a and end <= b for a, b in subject_spans):
                    continue
                cells[1] = cell[:start] + cell[end:]
                damaged_row = "|".join(cells)
                break
            assert damaged_row is not None, (subject, qualifier)
            assert damaged_row != row and subject in damaged_row, (subject, qualifier)
            _closure_refuses(
                lambda doc, old=row, new=damaged_row: doc.replace(old, new, 1),
                "so the exclusion it records has narrowed")


def test_221_the_closure_claims_the_unreachable_arms_were_induced() -> None:
    _closure_refuses(
        lambda doc: doc.replace(
            "**None of those was induced**",
            "Every failure mode was induced and nothing remains static-only", 1),
        "which no run established")


def test_222_run_id_exhaustion_returns_to_static_only() -> None:
    """`P6-RIDMAX` proved it on Windows; putting it back would understate the
    evidence as surely as over-claiming would overstate it."""
    _closure_refuses(
        lambda doc: doc.replace(
            "**No longer static-only:** cross-implementation parity (`P6-ORA`, D2b) and\nrun-ID exhaustion (`P6-RIDMAX`).",
            "| run-ID exhaustion (`P6-RIDMAX`) | not runnable. |", 1),
        "is back on the static-only list")


def test_223_cross_implementation_parity_returns_to_static_only() -> None:
    _closure_refuses(
        lambda doc: doc.replace(
            "**No longer static-only:** cross-implementation parity (`P6-ORA`, D2b) and\nrun-ID exhaustion (`P6-RIDMAX`).",
            "| cross-implementation parity (`P6-ORA`) | not runnable. |", 1),
        "is back on the static-only list")



# ===========================================================================
# R. The active wording after closure
# ===========================================================================
def test_224_the_scenario_banner_stays_at_run_4() -> None:
    """The banner described Run 4's world while Run 6 had closed the step, and
    called the corrections Run 6 exercised unexercised."""
    damaged = _swap(
        _PHASE6,
        "    WHAT HAS EXECUTED. Runs 1-6 have run.",
        "    WHAT HAS EXECUTED. Runs 1-4 have run, and the corrections after Run 4\n"
        "    are NOT among them - they are source.\n    Also,")
    _control("test_45", phase6=damaged)


def test_225_the_scenario_banner_loses_the_runtime_authority() -> None:
    """A banner edited after the run must not read as evidence for the tree that
    edited it."""
    damaged = _PHASE6.replace("a3924e0", "this commit")
    assert damaged != _PHASE6
    _control("test_45", phase6=damaged)


def test_226_the_scenario_banner_drops_the_static_only_boundary() -> None:
    damaged = _swap(
        _PHASE6,
        "    AND THE BOUNDARY SURVIVES THE PASS. An all-green run proves the scenarios\n",
        "    AND EVERYTHING IS NOW PROVEN. An all-green run proves the scenarios\n")
    damaged = damaged.replace("they remain static-only, and §5 of", "see §5 of")
    damaged = damaged.replace("were NOT induced and are not claimed -", "were covered -")
    _control("test_45", phase6=damaged)


def test_227_the_driver_banner_stays_at_run_4() -> None:
    """The active driver banner for a closed harness described it as still under
    runtime validation."""
    damaged = _swap(
        _HARNESS,
        "    PHASE 6 STEP 13 HAS COMPLETED WINDOWS/EXCEL RUNTIME VALIDATION.",
        "    The Phase-6 Step-13 block dot-sourced below is the part still under\n"
        "    runtime validation: the corrections have not yet been exercised on\n"
        "    Windows. Separately,")
    _control("test_75", harness=damaged)


def test_228_the_driver_banner_loses_the_closing_run() -> None:
    damaged = _HARNESS.replace("a3924e0", "a later commit")
    assert damaged != _HARNESS
    _control("test_75", harness=damaged)


def test_229_the_driver_banner_drops_the_runtime_authority_distinction() -> None:
    """Editing a banner does not make the commit that edited it runtime-proven."""
    damaged = _swap(
        _HARNESS,
        "    RUN 6 IS THE RUNTIME AUTHORITY, not the later documentation and control\n"
        "    commits that recorded it. Editing this banner does not make the commit that\n"
        "    edited it runtime-proven.\n\n",
        "")
    _control("test_75", harness=damaged)


def test_230_the_source_battery_docstring_stays_at_run_4() -> None:
    """The docstring said Runs 1-4 had executed, and test_45 REQUIRED it to."""
    conformance_path = Path(conformance.__file__)
    original = conformance_path.read_text(encoding="utf-8")
    damaged = original.replace(
        "WHERE STEP 13 ACTUALLY STANDS. Runs 1-6 have executed and Step 13 is CLOSED.",
        "WHERE STEP 13 ACTUALLY STANDS. Runs 1-4 have executed. What has NOT been\n"
        "exercised on Windows is everything settled after Run 4.", 1)
    assert damaged != original
    with tempfile.TemporaryDirectory(prefix="pccm-step13-docstring2-") as name:
        target = Path(name) / conformance_path.name
        target.write_text(damaged, encoding="utf-8")
        saved = conformance.__file__
        conformance.__file__ = str(target)
        try:
            refused = False
            try:
                conformance.test_45_the_current_wording_states_the_accepted_comparison_and_the_split()
            except AssertionError as error:
                refused = True
                assert "docstring" in str(error), error
        finally:
            conformance.__file__ = saved
    assert refused, "the module docstring may describe the pre-closure state"
