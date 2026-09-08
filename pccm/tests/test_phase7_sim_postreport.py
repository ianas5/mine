#!/usr/bin/env python3
"""PCCM Phase 7 Step-4 conformance tests for `src/vba/modSimPostReport.bas`.

The sensitivity pipeline: the CURRENT-run precondition, the resolved-model
bridge, the persisted TotalNom read, one-driver-at-a-time replay, the P7-2
kernel, ranking, the bounded `_SimData` block and its stamp.

--------------------------------------------------------------------------------
WHAT THESE TESTS PROVE, AND WHAT THEY DO NOT
--------------------------------------------------------------------------------
SOURCE CONFORMANCE AND ALGORITHMIC EQUIVALENCE, on Linux, now. The pipeline is
reproduced end to end against an INDEPENDENT Python reference built from the
accepted replay and the accepted kernel, and the module's ownership, ordering
and publication discipline are read from its source.

THE WORKSHEET HALF IS NOT EXECUTED. `modSimPostReport` reads and writes ranges,
and no VBA runtime or Excel exists here, so the transcriber cannot run it. What
is proved about the writes is their SHAPE and ORDER as written, and that is
stated rather than implied.

VBA EXECUTION CONFORMANCE IS NOT PROVED and is deferred to Phase-7 Windows
acceptance. Nothing here may be read as "VBA published a sensitivity table".

Runs standalone or under pytest.
"""

from __future__ import annotations

import re
import sys

import pytest
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder import load_sim_contract, load_structure_contract  # noqa: E402
from pccm_builder import sim_sensitivity as kernel  # noqa: E402
from pccm_builder.vba_source import VbaModule  # noqa: E402

import test_phase6_sim_engine_vba as engine  # noqa: E402
import test_phase7_sim_replay as replay  # noqa: E402
from phase6_vba_transcribe import _Ref, _val  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
POST_BAS = SRC_VBA / "modSimPostReport.bas"
SPEC = PCCM_ROOT / "spec"
SEED, N = 4242, 1000


def _module() -> VbaModule:
    return VbaModule(name="modSimPostReport", path=POST_BAS,
                     raw=POST_BAS.read_text(encoding="utf-8"))


def _code() -> str:
    return _module().code


def _procedure(name: str) -> str:
    code = _module().code_without_string_removal
    match = re.search(
        rf"^\s*(?:Public|Private)\s+(?:Function|Sub)\s+{re.escape(name)}\b", code, re.M)
    assert match, f"{name} is not declared"
    tail = code[match.start():]
    end = re.search(r"^\s*End\s+(?:Function|Sub)\s*$", tail, re.M)
    assert end, f"{name} has no End"
    return tail[: end.end()]


def _raw_contract() -> dict:
    return load_sim_contract(SPEC / "sim_contract.yaml").raw


# ---------------------------------------------------------------------------
# THE INDEPENDENT REFERENCE PIPELINE
# ---------------------------------------------------------------------------
# It uses the accepted replay and the accepted kernel - which is the point, they
# are the owners - but it composes them here rather than reading anything from
# modSimPostReport. What is being checked is the COMPOSITION: that the VBA does
# these steps, in this order, over all of them, once each.
def _pipeline(records):
    """(ranked, unranked) as the pipeline should produce them."""
    totals = replay._totals(records, seed=SEED, iterations=N)
    total_ranks = kernel.mid_ranks(totals)          # ONCE
    results = []
    for record in records:
        contributions = replay._replay(records, record["PermanentId"],
                                       seed=SEED, iterations=N)
        rho, status = kernel.rank_correlation(kernel.mid_ranks(contributions),
                                              total_ranks)
        results.append((record["PermanentId"], rho, status))
    order = kernel.rank_drivers(results)
    ranked = [(results[i][0], results[i][1]) for i in order]
    unranked = [r[0] for r in results if r[2] != kernel.SENSITIVITY_DEFINED]
    return ranked, unranked, results


def _cost(pid, dist="Triangular", lo=80.0, ml=100.0, hi=130.0, quantity=2.0):
    return engine._cost(pid, dist, lo, ml, hi, quantity=quantity)


def _risk(pid, dist="Triangular", lo=100.0, ml=200.0, hi=400.0, probability=0.3):
    return engine._risk(pid, dist, lo, ml, hi, probability=probability)


# ===========================================================================
# A. THE END-TO-END PIPELINE, against an independent composition
# ===========================================================================
def test_01_a_dominant_driver_ranks_first() -> None:
    """A cost line whose contribution dwarfs everything else must lead, and its
    rho must be near one because it very nearly IS the total."""
    records = replay._records(
        [_cost("C-001", "Uniform", 1000.0, None, 5000.0, quantity=10.0),
         _cost("C-002", "Uniform", 1.0, None, 2.0, quantity=1.0)],
        [_risk("R-001", "Uniform", 1.0, None, 3.0, probability=0.5)])
    ranked, unranked, _ = _pipeline(records)
    assert ranked[0][0] == "C-001", ranked
    assert ranked[0][1] > 0.95, ranked[0]
    assert unranked == []


def test_02_a_risk_at_twenty_percent_matches_an_independent_mid_rank_pearson() -> None:
    """The large tie block, checked against a reference written from the
    definition rather than against the same code under another name."""
    import math
    records = replay._records([], [_risk("R-001", "Beta-PERT", 100.0, 200.0, 400.0,
                                         probability=0.2)])
    totals = replay._totals(records, seed=SEED, iterations=N)
    contributions = replay._replay(records, "R-001", seed=SEED, iterations=N)
    zeros = contributions.count(0.0)
    assert 0.7 * N < zeros < 0.9 * N, zeros

    x, y = kernel.mid_ranks(contributions), kernel.mid_ranks(totals)
    mx, my = sum(x) / N, sum(y) / N
    expected = (sum((a - mx) * (b - my) for a, b in zip(x, y))
                / math.sqrt(sum((a - mx) ** 2 for a in x) * sum((b - my) ** 2 for b in y)))
    ranked, _unranked, _ = _pipeline(records)
    assert math.isclose(ranked[0][1], expected, rel_tol=1e-12, abs_tol=1e-12)


def test_03_a_zero_variance_driver_is_reported_but_not_ranked() -> None:
    records = replay._records(
        [_cost("C-001", "Uniform", 10.0, None, 20.0),
         _cost("C-002", "Uniform", 50.0, None, 50.0)],   # degenerate: no variance
        [_risk("R-001", probability=0.0)])               # never occurs: no variance
    ranked, unranked, results = _pipeline(records)
    assert [pid for pid, _ in ranked] == ["C-001"], ranked
    assert sorted(unranked) == ["C-002", "R-001"]
    # REPORTED, not deleted: every driver still has a record.
    assert len(results) == 3


def test_04_positive_and_negative_drivers_both_appear_with_their_sign() -> None:
    records = replay._records(
        [_cost("C-001", "Uniform", 100.0, None, 900.0, quantity=1.0),
         _cost("C-002", "Uniform", 100.0, None, 900.0, quantity=1.0)], [])
    ranked, _unranked, results = _pipeline(records)
    assert len(ranked) == 2
    # Two independent drivers of one total: both correlate POSITIVELY with it,
    # and the signs are retained rather than discarded by the ordering.
    assert all(rho > 0.0 for _pid, rho in ranked), ranked
    assert all(abs(a) >= abs(b) for a, b in zip([r for _, r in ranked],
                                                [r for _, r in ranked][1:]))


def test_05_equal_magnitude_drivers_keep_the_permanent_id_tie_break() -> None:
    """Carried all the way through the pipeline, not just inside the kernel."""
    results = [("C-010", 0.4, kernel.SENSITIVITY_DEFINED),
               ("C-002", -0.4, kernel.SENSITIVITY_DEFINED),
               ("R-001", 0.4, kernel.SENSITIVITY_DEFINED)]
    order = kernel.rank_drivers(results)
    assert [results[i][0] for i in order] == ["C-002", "C-010", "R-001"]


# ===========================================================================
# B. THE TOTAL'S RANKS ARE COMPUTED ONCE
# ===========================================================================
def test_06_the_total_is_ranked_once_and_reused_for_every_driver() -> None:
    body = _procedure("RunSensitivity")
    assert body.count("SimSensitivityMidRanks(") == 1, (
        "the total is ranked more than once, or not in the orchestrator")
    # AND THE PER-DRIVER CALL TAKES THE ALREADY-RANKED VECTOR.
    analyse = _procedure("AnalyseDrivers")
    assert "SimSensitivityMidRanks(" not in analyse, (
        "the total is re-ranked inside the per-driver loop")
    assert "SimSensitivitySpearman(" in analyse
    assert "totalRanks" in analyse


def test_07_every_driver_is_processed_exactly_once() -> None:
    analyse = _procedure("AnalyseDrivers")
    assert re.search(r"For index = 0 To driverCount - 1", analyse), (
        "the analysis does not walk every driver")
    assert analyse.count("SimEngineReplayDriver(") == 1
    assert analyse.count("SimSensitivitySpearman(") == 1
    # BY PERMANENT ID, AND BY THE LOOP'S OWN INDEX. `drivers(LBound(drivers))`
    # also ends in `.PermanentId` and would replay one driver D times.
    assert "drivers(LBound(drivers) + index).PermanentId, contributions" in analyse, (
        "the replayed driver is not the one this iteration of the loop reached")
    assert "results(index).PermanentId = drivers(LBound(drivers) + index).PermanentId" in analyse


# ===========================================================================
# C. MEMORY - O(N), never O(D x N)
# ===========================================================================
def test_08_one_contribution_vector_exists_at_a_time() -> None:
    """The 240 MB matrix is not built here either. `contributions` is a single
    vector that the next driver overwrites; nothing indexes it by driver."""
    analyse = _procedure("AnalyseDrivers")
    assert "Dim contributions() As Double" in analyse
    # NO DRIVER INDEX ON IT. `contributions(index)` or a second dimension would
    # be the matrix arriving by the back door.
    assert not re.search(r"contributions\s*\(\s*index", analyse)
    assert not re.search(r"ReDim\s+contributions\s*\([^)]*,", analyse)
    code = _code()
    assert not re.search(r"ReDim\s+\w+\s*\(0 To driverCount - 1, 0 To", code), (
        "a driver x iteration container is allocated")


def test_09_nothing_retains_a_per_driver_sample_or_a_matrix() -> None:
    code = _code()
    for banned in ("samples", "Samples", "matrix", "Matrix", "AnnualRecords"):
        assert banned not in code, banned
    # The only D-sized things are the RESULT records and the order permutation.
    sized = set(re.findall(r"ReDim (\w+)\(0 To driverCount - 1\)", code))
    assert sized <= {"results"}, sorted(sized)


# ===========================================================================
# D. OWNERSHIP - it orchestrates and computes nothing
# ===========================================================================
def test_10_no_mathematics_is_reimplemented_here() -> None:
    code = _code()
    for banned in ("SafeProduct", "SafeSignedSum", "midrank", "MidRank(",
                   "Pearson", "SimRngNext", "SimSampleUniform", "SimSampleTriangular",
                   "SimSamplePreparedBeta", "SimSampleBernoulli", "SimRngJump"):
        assert banned not in code, f"modSimPostReport reimplements {banned!r}"
    # THE IDENTITY IS READ, NEVER DERIVED. A token search for "Fingerprint("
    # would also flag the accepted accessor, which is the correct behaviour -
    # so the rule names WHERE those two values may come from instead.
    for kind in ("Fingerprint", "Digest"):
        uses = [line.strip() for line in code.splitlines()
                if re.search(rf"\b\w*{kind}\w*\s*\(", line)]
        for line in uses:
            assert f"modSimReport.PCCM_Simulation" in line, (
                f"a {kind.lower()} is produced here rather than read: {line}")


def test_11_each_step_is_delegated_to_its_accepted_owner() -> None:
    code = _code()
    assert "modSimReport.PCCM_SimulationStatus()" in code, "state is re-derived"
    assert "modCalcReport.CalcPrepareSimulationInputs(" in code, "the model is re-resolved"
    assert "modSimEngine.SimEngineReplayDriver(" in code, "replay is re-implemented"
    for owned in ("SimSensitivityMidRanks", "SimSensitivitySpearman", "SimSensitivityRank"):
        assert f"modSimSensitivity.{owned}(" in code, owned


def test_12_it_starts_no_simulation_and_consumes_no_run_identity() -> None:
    code = _code()
    for banned in ("PCCM_RunSimulation", "SimEngineRun(", "AllocateAutoNonce",
                   "SIM_PENDING_AUTO_NONCE_CELL", "NextAutoNonce", "CandidateRunId",
                   "WriteAttemptBlock", "WriteStatusBlock", "FinalCommit"):
        assert banned not in code, f"sensitivity reaches {banned!r}"


def test_13_the_simulation_does_not_run_sensitivity_for_you() -> None:
    """Explicit post-processing. A successful run must stay successful even if
    the analysis of it later fails."""
    report = (SRC_VBA / "modSimReport.bas").read_text(encoding="utf-8")
    assert "modSimPostReport" not in report
    assert "RunSensitivity" not in report


def test_14_it_writes_no_iteration_row_and_no_result_digest() -> None:
    code = _code()
    assert "SIM_ITER_A_TOTAL_NOMINAL_COLUMN" in code, "it must READ the totals"
    # ...and only read them.
    publish = _procedure("Publish")
    for banned in ("SIM_ITER_", "SIM_IDENTITY_ROW_RESULT_DIGEST",
                   "SIM_IDENTITY_ROW_RUN_ID", "SIM_SUMMARY_", "SIM_CONTINGENCY_"):
        assert banned not in publish, f"publication writes {banned!r}"


# ===========================================================================
# E. THE PRECONDITION AND REFUSALS
# ===========================================================================
def test_15_only_a_current_run_may_be_analysed() -> None:
    body = _procedure("RequireCurrentRun")
    assert "PCCM_SimulationStatus()" in body
    assert "SIM_STATE_CURRENT" in body
    # STALE and INVALID both fall through the same inequality; naming only one
    # would leave the other admitted.
    assert re.search(r"StrComp\(status, SIM_STATE_CURRENT, vbBinaryCompare\) <> 0", body)
    # AND BOTH GUARDS ARE REAL CONDITIONS. A message left behind an `If False`
    # is not a refusal, so the tested expression is read rather than the text
    # that follows it.
    assert re.search(r"If Len\(status\) = 0 Then", body), (
        "the no-successful-run guard is not a test of the status")
    assert "no successful simulation" in body
    assert "If False Then" not in body, "a precondition was short-circuited"


def test_16_a_refusal_writes_nothing_at_all() -> None:
    """Every refusal path exits before Publish is reached."""
    run = _procedure("RunSensitivity")
    refusals = run.count("RunSensitivity = Refused(detail)")
    assert refusals >= 6, refusals
    # EVERY PRECONDITION REFUSAL RETURNS BEFORE PUBLICATION. The last refusal
    # is Publish's own - a write that fails is still a refusal - so the rule is
    # that every OTHER one precedes it, not that none follows.
    at_publish = run.index("Publish(")
    before = [m.start() for m in re.finditer(r"RunSensitivity = Refused\(detail\)", run)]
    assert sum(1 for position in before if position < at_publish) == len(before) - 1, (
        "a refusal path can be reached after publication has begun")
    refused = _procedure("Refused")
    for banned in ("Range", "Value2", "ClearContents", "Stamp"):
        assert banned not in refused, f"a refusal touches {banned!r}"


# ===========================================================================
# F. PUBLICATION SAFETY
# ===========================================================================
def test_17_the_published_marker_is_cleared_first_and_written_last() -> None:
    """A block that fails part way through must not carry a current stamp."""
    publish = _procedure("Publish")
    first = publish.index("SIM_SENSITIVITY_STAMP_ROW_PUBLISHED")
    last = publish.rindex("SIM_SENSITIVITY_STAMP_ROW_PUBLISHED")
    assert first != last, "the marker is written only once"
    assert "vbNullString" in publish[first:first + 120], (
        "the marker is not CLEARED before the block is rewritten")
    assert "SIM_SENSITIVITY_PUBLISHED" in publish[last:last + 120], (
        "the marker is not SET at the end")
    # NOTHING FOLLOWS IT but the success return.
    tail = publish[last:]
    assert tail.count("Value2 =") == 1, "a write follows the published marker"


def test_18_the_identity_is_written_before_the_marker() -> None:
    publish = _procedure("Publish")
    marker = publish.rindex("SIM_SENSITIVITY_STAMP_ROW_PUBLISHED")
    for field in ("RUN_ID", "EFFECTIVE_SEED", "REQUEST_FINGERPRINT",
                  "RESULT_DIGEST", "ITERATIONS", "RECORD_COUNT"):
        at = publish.index(f"SIM_SENSITIVITY_STAMP_ROW_{field}")
        assert at < marker, f"{field} is stamped after the published marker"


def test_19_the_whole_result_is_built_before_anything_is_written() -> None:
    run = _procedure("RunSensitivity")
    for step in ("RequireCurrentRun(", "ResolveDrivers(", "ReadTotals(",
                 "SimSensitivityMidRanks(", "AnalyseDrivers(", "SimSensitivityRank("):
        assert run.index(step) < run.index("Publish("), f"{step} runs after publication"
    assert "NOTHING HAS BEEN WRITTEN UNTIL HERE" in _module().raw


def test_20_surplus_rows_from_a_larger_previous_result_are_cleared() -> None:
    """A later model can have fewer drivers than the bank already holds.
    Overwriting the first n would leave the remainder visible and
    indistinguishable from the new result."""
    publish = _procedure("Publish")
    assert publish.index("ClearRecords") < publish.index("Value2 = block"), (
        "the block is written before the old rows are cleared")
    clear = _procedure("ClearRecords")
    assert "ClearContents" in clear
    # TO THE CEILING, not to the new count: the surplus is what must go.
    assert "SIM_MAX_ITERATIONS" in clear
    assert "driverCount" not in clear


def test_21_the_record_count_bounds_what_is_authoritative() -> None:
    publish = _procedure("Publish")
    assert "SIM_SENSITIVITY_STAMP_ROW_RECORD_COUNT).Value2 = driverCount" in publish
    # AND THE BUILT BLOCK IS EXACTLY THAT MANY ROWS.
    assert "ReDim block(1 To driverCount, 1 To SIM_SENSITIVITY_FIELD_COUNT)" in publish
    assert "If slot <> driverCount Then" in publish, (
        "nothing checks that every driver produced exactly one record")


# ===========================================================================
# G. THE RECORD SHAPE
# ===========================================================================
def test_22_every_contracted_field_is_written_and_no_other() -> None:
    fill = _procedure("FillRecord")
    contracted = [c["key"] for c in
                  _raw_contract()["sim_data"]["sensitivity_records"]["columns"]]
    assert contracted == ["driver_id", "driver_type", "driver_name", "rho",
                          "abs_rho", "rank", "direction", "status"]
    for key in contracted:
        assert f"SIM_SENSITIVITY_OFFSET_{key.upper()}" in fill, key
    written = set(re.findall(r"SIM_SENSITIVITY_OFFSET_(\w+)", fill))
    assert written == {key.upper() for key in contracted}, sorted(written)
    # THE IDENTITY FIELDS ARE COMMON; THE MEASURE FIELDS ARE PER-ARM. A measure
    # written only when the driver has a rho would leave a stale value in the
    # cell of a driver that does not, so the two branches are read separately -
    # and the three identity fields are asserted to sit above the branch, where
    # both arms get them.
    identity = ("driver_id", "driver_type", "driver_name")
    measures = ("rho", "abs_rho", "rank", "direction", "status")
    assert set(identity) | set(measures) == set(contracted)
    branch = fill.index("    If record.Status = SIM_SENSITIVITY_DEFINED Then")
    common, rest = fill[:branch], fill[branch:]
    defined = rest[:rest.index("    Else")]
    undefined = rest[rest.index("    Else"):]
    for key in identity:
        assert f"SIM_SENSITIVITY_OFFSET_{key.upper()} + 1)" in common, (
            f"{key} is written inside a branch and would be missing from the other")
    for key in measures:
        token = f"SIM_SENSITIVITY_OFFSET_{key.upper()} + 1)"
        assert token in defined, f"{key} is not written for a ranked driver"
        assert token in undefined, f"{key} is not written for a zero-variance driver"


def test_23_a_zero_variance_record_carries_the_label_and_no_rho() -> None:
    fill = _procedure("FillRecord")
    undefined = fill[fill.index("Else"):]
    assert "SENSITIVITY_NO_VARIANCE_LABEL" in undefined
    # NO RHO, NO RANK, NO DIRECTION - printing 0 would say a relationship was
    # looked for and not found.
    for offset in ("RHO", "ABS_RHO", "RANK", "DIRECTION"):
        assert f"SIM_SENSITIVITY_OFFSET_{offset} + 1) = vbNullString" in undefined, offset
    raw = _module().raw
    assert 'SENSITIVITY_NO_VARIANCE_LABEL As String = "n/a - no variance"' in raw


def test_24_the_ranked_records_come_first_and_carry_their_rank() -> None:
    publish = _procedure("Publish")
    ranked_at = publish.index("For position = 0 To eligibleCount - 1")
    rest_at = publish.index("If results(index).Status <> SIM_SENSITIVITY_DEFINED")
    assert ranked_at < rest_at, "the unranked records are written first"
    assert "FillRecord block, slot + 1, results(index), position + 1" in publish
    assert "FillRecord block, slot + 1, results(index), 0" in publish


def test_25_no_variance_share_or_squared_rho_is_produced() -> None:
    code = _code()
    for banned in ("rho * rho", "rho ^ 2", "Squared", "percent", "Percent",
                   "Contribution %", "variance share"):
        assert banned not in code, banned


def test_26_no_top_n_truncation_and_no_subsampling() -> None:
    code = _code()
    for banned in ("TopN", "Top_N", "topN", "Subsample", "subsample", "SampleEvery"):
        assert banned not in code, banned
    # Full N: the iteration count comes from the published run and is not capped.
    assert "run.Iterations" in code
    assert not re.search(r"Iterations\s*=\s*\d+", code), "an iteration literal is imposed"
    # THE POPULATION IS NOT NARROWED. `eligibleCount` and `driverCount` are
    # produced by the kernel and the model; a reassignment of either here is a
    # truncation whatever it is called.
    for counter in ("eligibleCount", "driverCount"):
        # ANYWHERE ON THE LINE. `If eligibleCount > 10 Then eligibleCount = 10`
        # is a truncation that a line-start pattern would walk straight past.
        assignments = re.findall(rf"{counter}\s*=(?!=)\s*(.+)", code)
        assert assignments == [], f"{counter} is reassigned: {assignments}"


# ===========================================================================
# H. REGISTRATION AND THE ENDPOINT
# ===========================================================================
def test_27_the_module_and_endpoint_are_declared() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    declared = {m.name: m for m in structure.vba_modules}
    assert "modSimPostReport" in declared
    entry = declared["modSimPostReport"]
    assert entry.generated is False
    for phrase in ("CURRENT-run precondition", "published last", "owns no RNG"):
        assert phrase in entry.responsibility, phrase
    assert _module().public_procedures == ["PCCM_RunSensitivity"], (
        _module().public_procedures)


def test_28_the_sensitivity_kernel_and_the_engine_stay_where_they_are() -> None:
    """P7-4 moved no mathematics into the orchestrator and none out of it."""
    for name, banned in (
        ("modSimSensitivity", ("Range", "Worksheet", "SimEngine", "SimRng", "_SimData")),
        ("modSimEngine", ("Range", "Worksheet", "_SimData", "modSimPostReport")),
    ):
        path = SRC_VBA / f"{name}.bas"
        code = VbaModule(name=name, path=path,
                         raw=path.read_text(encoding="utf-8")).code
        for token in banned:
            assert token not in code, f"{name} acquired {token!r}"


# ===========================================================================
# I. THE SENSITIVITY SHEET - lookups, bounds, and whose answer it is
# ===========================================================================
def _shell() -> dict:
    from pccm_builder.spec_loader import load_spec
    return load_spec(SPEC / "workbook.yaml").phase6_shell["sensitivity"]


def _built_sheet():
    import openpyxl
    workbook = openpyxl.load_workbook(PCCM_ROOT / "build" / "PCCM_stageA.xlsx")
    return workbook[_shell()["sheet"]]


def test_29_the_table_shows_the_eight_contracted_fields() -> None:
    shell, sheet = _shell(), _built_sheet()
    contracted = [c["header"] for c in
                  _raw_contract()["sim_data"]["sensitivity_records"]["columns"]]
    headers = [sheet[f"{c['column']}{shell['header_row']}"].value
               for c in shell["columns"]]
    assert headers == contracted == [
        "Driver ID", "Type", "Name", "Rho", "|Rho|", "Rank", "Direction", "Status"]


def test_30_every_table_cell_is_a_lookup_and_computes_nothing() -> None:
    """The sheet SHOWS the authoritative answer; it does not reproduce it."""
    shell, sheet = _shell(), _built_sheet()
    forbidden = ("RANK(", "CORREL(", "PEARSON(", "AVERAGE(", "STDEV", "SUMPRODUCT(",
                 "PERCENTILE", "RAND(", "OFFSET(", "INDIRECT(", "LARGE(", "SMALL(",
                 "SORT(", "^2")
    seen = 0
    for row_index in range(int(shell["row_window"])):
        row = int(shell["first_row"]) + row_index
        for column in shell["columns"]:
            formula = sheet[f"{column['column']}{row}"].value
            assert isinstance(formula, str) and formula.startswith("="), (column, row)
            for banned in forbidden:
                assert banned not in formula.upper(), (banned, column["key"])
            assert "_SimData!" in formula, "a table cell does not read the persisted block"
            seen += 1
    assert seen == int(shell["row_window"]) * 8, seen


def test_31_a_row_beyond_the_persisted_count_is_blank() -> None:
    """A later model with fewer drivers cannot leave a previous run's surplus
    rows on display: the count decides, not the window."""
    shell, sheet = _shell(), _built_sheet()
    stamp = _raw_contract()["sim_data"]["sensitivity_records"]["stamp"]
    count_row = next(f["row"] for f in stamp["fields"] if f["key"] == "record_count")
    for row_index in (0, 7, 199):
        formula = sheet[f"{shell['columns'][0]['column']}"
                        f"{int(shell['first_row']) + row_index}"].value
        assert f"IF({row_index + 1}>IF(" in formula, (row_index, formula[:80])
        # THE STAMP COLUMNS COME FROM THE CONTRACT. Spelling them here as $J$
        # and $S$ is what let the P7-1 allocation survive review: two literals
        # agreeing with a wrong contract look exactly like a checked copy
        # agreeing with a right one.
        columns = stamp["bank_value_columns"]
        for bank in ("A", "B"):
            assert f"${columns[bank]}${count_row}" in formula, (bank, formula)


def test_32_an_unpublished_block_shows_nothing_at_all() -> None:
    shell, sheet = _shell(), _built_sheet()
    stamp = _raw_contract()["sim_data"]["sensitivity_records"]["stamp"]
    published_row = next(f["row"] for f in stamp["fields"] if f["key"] == "published")
    formula = sheet[f"{shell['columns'][0]['column']}{shell['first_row']}"].value
    columns = stamp["bank_value_columns"]
    assert (f'${columns["A"]}${published_row},_SimData!${columns["B"]}'
            f'${published_row})<>"PUBLISHED",""') in formula, formula


def test_33_the_sheet_says_whose_answer_it_is() -> None:
    """THE CASE THAT MUST NEVER BE SILENT. A table produced for run A, shown
    while run B is published, has to say so."""
    shell, sheet = _shell(), _built_sheet()
    availability = sheet[f"{shell['columns'][1]['column']}{shell['availability_row']}"].value
    assert "No simulation has been published." in availability
    assert "Not produced for this run." in availability
    assert "CURRENT for run " in availability
    assert "NOT CURRENT - this table belongs to run " in availability
    # It decides by comparing the STAMP against the published identity, not by
    # a flag someone could set.
    #
    # THE ADDRESSES ARE DERIVED, NOT TYPED. This assertion used to name $J$10
    # and $J$11 literally. When the P7-4 runtime correction moved the
    # sensitivity block off J/S onto CC/CL, the formula in workbook.yaml stayed
    # behind - and so did this test, which went on passing because it had
    # restated the same stale address the formula had. A test that agrees with a
    # defect is worse than no test: it certifies it.
    raw = _raw_contract()["sim_data"]
    stamp = raw["sensitivity_records"]["stamp"]
    stamp_rows = {f["key"]: f["row"] for f in stamp["fields"]}
    identity_rows = {f["key"]: f["row"] for f in raw["run_identity"]["fields"]}
    for key, label in (("request_fingerprint", "fingerprint"),
                       ("result_digest", "digest")):
        for bank, column in stamp["bank_value_columns"].items():
            assert f"${column}${stamp_rows[key]}" in availability, (
                f"bank {bank}'s stamped {label} is never read"
            )
        for bank, column in raw["run_identity"]["bank_value_columns"].items():
            assert f"${column}${identity_rows[key]}" in availability, (
                f"bank {bank}'s published {label} is never compared against"
            )


def test_34_the_sheet_carries_no_tornado_and_no_variance_share() -> None:
    """Phase 8 owns the chart. Phase 7 owns the table it will read."""
    import openpyxl
    workbook = openpyxl.load_workbook(PCCM_ROOT / "build" / "PCCM_stageA.xlsx")
    sheet = workbook[_shell()["sheet"]]
    text = " ".join(str(cell.value) for row in sheet.iter_rows() for cell in row
                    if cell.value is not None)
    for banned in ("Tornado", "tornado", "% of variance", "variance share", "R²", "rho²"):
        assert banned not in text, banned
    assert not getattr(sheet, "_charts", []), "a chart was placed on the Sensitivity sheet"


def test_35_the_display_window_cannot_silently_exceed_the_block() -> None:
    """The window is display; the persisted block is authority. A window taller
    than the sheet's own record area would read rows the block never owns."""
    shell = _shell()
    records = _raw_contract()["sim_data"]["sensitivity_records"]
    assert int(shell["first_row"]) > int(shell["header_row"])
    assert int(shell["row_window"]) >= 1
    # The persisted block starts at the shared first record row and runs to the
    # technical ceiling, so any window up to that is readable.
    ceiling = _raw_contract()["iterations"]["technical_ceiling"]["max_iterations_representable"]
    assert int(shell["row_window"]) <= ceiling
    assert int(records["first_record_row"]) == 34


# ===========================================================================
# THE 300-DRIVER TRUNCATION, AND WHY test_35 DID NOT CATCH IT
# ===========================================================================
# The design-scale timing probe persisted 300 sensitivity records and the sheet
# displayed 200 of them, silently. `test_35` passed throughout: it asked whether
# the window was within the block's physical reach, which it was, and never
# asked the question that mattered - whether a reader could tell that rows were
# missing. A control that pins a BOUND rather than the PROPERTY the bound exists
# to protect is how a defect reaches Windows.
#
# These three ask the property instead, and they fail on the pre-correction
# workbook: at row_window 200 the window does not clear the measured population
# and the availability line carries no disclosure at all.
DESIGN_SCALE_DRIVERS = 300  # the measured design-scale probe: 300 drivers, 10k iterations


def test_36_the_window_clears_the_largest_measured_model() -> None:
    """A window smaller than a model the project has actually run is a Top-N
    truncation with a different name, and `ranking.top_n_truncation` is false."""
    shell = _shell()
    ranking = _raw_contract()["sensitivity"]["ranking"]
    assert ranking["top_n_truncation"] is False
    assert ranking["population"] == "every eligible non_zero_variance driver"
    assert int(shell["row_window"]) > DESIGN_SCALE_DRIVERS, (
        f"the sheet shows {shell['row_window']} rows but the design-scale probe "
        f"produced {DESIGN_SCALE_DRIVERS} ranked drivers; the materialisation "
        "must not be the thing that truncates a result the contract says is complete"
    )


def test_37_an_overflow_announces_itself_on_the_sheet() -> None:
    """Capacity alone cannot be proved sufficient - no contract names a maximum
    driver population - so the sheet must say when it has run out of room.

    The window number in the formula is a CHECKED COPY of `row_window`, not a
    second declaration: a raised window with a stale threshold in the formula
    would disclose at the wrong count, or never.
    """
    shell = _shell()
    display = _raw_contract()["sensitivity"]["display"]
    assert display["silent_truncation_permitted"] is False
    assert display["overflow_disclosure_required"] is True
    assert display["sheet_window_owner"] == (
        "workbook.yaml: phase6_shell.sensitivity.row_window")

    window = int(shell["row_window"])
    formula = _built_sheet()[
        f"{shell['columns'][1]['column']}{shell['availability_row']}"].value
    assert f">{window}" in formula, (
        "the availability line must compare the persisted record count against "
        f"the window it actually renders ({window})"
    )
    assert "showing the first" in formula and "ranked drivers" in formula, (
        "an overflow must be announced in words the reader can act on"
    )
    assert "the complete result is persisted" in formula, (
        "the reader must be told the missing rows exist, not merely that they "
        "are not shown"
    )
    # BOTH banks, because the answer is whichever bank is active.
    records = _raw_contract()["sim_data"]["sensitivity_records"]
    count_row = next(f["row"] for f in records["stamp"]["fields"]
                     if f["key"] == "record_count")
    for bank, column in records["stamp"]["bank_value_columns"].items():
        assert f"${column}${count_row}" in formula, (
            f"the disclosure never reads bank {bank}'s record count"
        )


def test_38_the_disclosure_did_not_disturb_the_four_existing_answers() -> None:
    """The correction appends; it does not rewrite what was accepted.

    And it appends only where rows are actually shown: the two states that
    display nothing at all - no publication, and not produced for this run -
    cannot acquire a note about rows they are not displaying.
    """
    shell = _shell()
    formula = _built_sheet()[
        f"{shell['columns'][1]['column']}{shell['availability_row']}"].value
    for answer in ("No simulation has been published.",
                   "Not produced for this run. Run Sensitivity to produce it.",
                   "CURRENT for run ",
                   "NOT CURRENT - this table belongs to run "):
        assert answer in formula, f"the accepted answer {answer!r} was lost"
    head, _, tail = formula.partition("showing the first")
    assert "No simulation has been published." in head
    assert "Not produced for this run." in head
    assert tail, "the disclosure is missing"
    # The disclosure sits after the CURRENT/NOT CURRENT pair, so it can only
    # extend an answer that is already showing rows.
    assert head.rindex("NOT CURRENT - this table belongs to run ") < len(head)


# ===========================================================================
# H. THE DRIVER LABEL - P8-3 WINDOWS RUN 4
# ===========================================================================
# WHAT THE LIVE RUN FOUND. `PCCM_RunSensitivity` published "5 ranked of 5
# drivers" and two of the five Name cells held a NUMERIC ZERO. The W4 fixture
# carries three cost lines and two risks, and the two zeroes were the two risks.
#
# THE DEFECT, FROM THE CONTRACT AND NOT FROM THE COINCIDENCE. Each register
# carries its label under a different key, and driver_contract.yaml says which:
#
#   cost_lines.description    required: true    the cost line's only label
#   risk_register.risk_name   required: true    the risk's user-facing name
#   risk_register.description required: FALSE   an optional free-text note
#
# DriverNameOf read DESCRIPTION for BOTH. For a cost line that is the required
# label and was always right. For a risk it is the optional note, so a risk
# without one published no name at all - and the empty string that produced
# reached the sheet through a `.Value2` array write, which stored it as a
# numeric zero rather than leaving the cell empty. That is the `0` Run 4 saw.
#
# THESE TESTS EXECUTE THE REAL SOURCE. The transcriber compiles DriverNameOf out
# of the .bas file and the two registers are simulated, so the assertions are
# about what the module DOES, not about what its text looks like - and they fail
# against the pre-fix source, which is the only way to know they test anything.
def _register_columns() -> dict[str, int]:
    """The register column ordinals, READ OUT OF THE GENERATED modConstants.

    Not retyped here and not taken from the manifest: modConstants is what the
    module under test actually compiles against, so a column that moved would
    move here too rather than leaving this fixture describing a workbook that
    no longer exists."""
    generated = (PCCM_ROOT / "build" / "vba" / "modConstants.bas").read_text(
        encoding="utf-8")
    found = dict(re.findall(
        r"^Public Const (COL_(?:RISK_REGISTER|COST_LINES)_\w+) As Long = (\d+)",
        generated, re.M))
    assert found, "modConstants publishes no register column ordinals"
    return {name: int(value) for name, value in found.items()}


def _register_fixture(risk_name: str, risk_description: str,
                      cost_description: str) -> dict:
    """Two registers, keyed the way modDrivers keys them. Column ordinals are
    the generated modConstants values, never retyped here."""
    constants = _register_columns()
    return {
        "RISK": {
            "R-001": {
                constants["COL_RISK_REGISTER_RISK_ID"]: "R-001",
                constants["COL_RISK_REGISTER_RISK_NAME"]: risk_name,
                constants["COL_RISK_REGISTER_DESCRIPTION"]: risk_description,
            },
        },
        "COST": {
            "CL-001": {
                constants["COL_COST_LINES_COST_LINE_ID"]: "CL-001",
                constants["COL_COST_LINES_CATEGORY"]: "Civils",
                constants["COL_COST_LINES_DESCRIPTION"]: cost_description,
            },
        },
    }


def _published_name(permanent_id: str, registers: dict) -> str:
    """Run the module's own DriverNameOf over a simulated workbook."""
    from phase6_vba_transcribe import build as _build

    seen: dict = {}

    def row_of_id(kind, driver_id):
        table = registers.get(_val(kind), {})
        ids = list(table)
        value = _val(driver_id)
        return ids.index(value) + 1 if value in ids else 0

    def register_table(kind):
        return _val(kind)

    def cell_in(table, row, column):
        rows = registers.get(_val(table), {})
        key = list(rows)[_val(row) - 1]
        # A CELL THAT WAS NEVER WRITTEN IS EMPTY, exactly as an unpopulated
        # optional column is on the sheet.
        return rows[key].get(_val(column), None)

    def text_of(cell):
        value = _val(cell)
        # modWorkbook.TextOf is Trim$(CStr(Target.Value & "")): an empty cell
        # becomes the empty string, which is what then reached .Value2.
        return "" if value is None else str(value).strip()

    namespace = _build(
        {"modSimPostReport": POST_BAS},
        dict(engine._constants(), **_register_columns()),
        only={"modSimPostReport": {"DriverNameOf"}},
        extra={
            "RiskKind": lambda: "RISK",
            "CostKind": lambda: "COST",
            "RowOfId": row_of_id,
            "RegisterTable": register_table,
            "CellIn": cell_in,
            "TextOf": text_of,
            "seen": seen,
        })
    return namespace["DriverNameOf"](_Ref(permanent_id))


def test_40_a_risk_publishes_its_risk_name_when_the_description_is_blank() -> None:
    """THE EXACT DEFECT. This is the W4 shape: Risk Name populated, Description
    never written. Against the pre-fix source it publishes the empty string."""
    registers = _register_fixture(risk_name="GateB R-001", risk_description="",
                                  cost_description="GateB CL-001")
    assert _published_name("R-001", registers) == "GateB R-001"


def test_41_a_risk_publishes_the_name_even_when_both_columns_are_populated() -> None:
    """THE DISCRIMINATOR. With two DIFFERENT texts present, only the field that
    is actually read can be observed - a blank description alone could be
    satisfied by any fallback rule."""
    registers = _register_fixture(risk_name="THE RISK NAME",
                                  risk_description="THE DESCRIPTION NOTE",
                                  cost_description="GateB CL-001")
    published = _published_name("R-001", registers)
    assert published == "THE RISK NAME"
    assert published != "THE DESCRIPTION NOTE", (
        "the risk still publishes its optional description")


def test_42_a_risk_never_publishes_its_identifier_as_its_name() -> None:
    """AND NOT THE ID EITHER. The identity field is published separately; a name
    that repeated it would lose the label without looking empty."""
    registers = _register_fixture(risk_name="THE RISK NAME",
                                  risk_description="THE DESCRIPTION NOTE",
                                  cost_description="GateB CL-001")
    assert _published_name("R-001", registers) != "R-001"


def test_43_the_cost_line_label_is_unchanged() -> None:
    """THE COST-LINE SIDE WAS ALWAYS RIGHT and must stay exactly as it was: its
    description is the `required: true` column that labels it, and it has no
    separate name field to move to."""
    registers = _register_fixture(risk_name="THE RISK NAME",
                                  risk_description="THE DESCRIPTION NOTE",
                                  cost_description="GateB CL-001")
    assert _published_name("CL-001", registers) == "GateB CL-001"
    # AND IT IS THE DESCRIPTION IT READS, proved by moving that value alone.
    moved = _register_fixture(risk_name="THE RISK NAME",
                              risk_description="THE DESCRIPTION NOTE",
                              cost_description="A DIFFERENT COST LABEL")
    assert _published_name("CL-001", moved) == "A DIFFERENT COST LABEL"


def test_44_every_published_name_is_text_and_none_is_a_numeric_zero() -> None:
    """THE W4 POPULATION, AS THE FIXTURE ACTUALLY WRITES IT: three cost lines
    with descriptions, two risks with names and NO description. Run 4 published
    two numeric zeroes here."""
    constants = _register_columns()
    registers = {"RISK": {}, "COST": {}}
    for index in (1, 2):
        registers["RISK"][f"R-00{index}"] = {
            constants["COL_RISK_REGISTER_RISK_ID"]: f"R-00{index}",
            constants["COL_RISK_REGISTER_RISK_NAME"]: f"GateB R-00{index}",
            # THE FIXTURE NEVER WRITES THIS ONE.
        }
    for index in (1, 2, 3):
        registers["COST"][f"CL-00{index}"] = {
            constants["COL_COST_LINES_COST_LINE_ID"]: f"CL-00{index}",
            constants["COL_COST_LINES_DESCRIPTION"]: f"GateB CL-00{index}",
        }
    published = {driver: _published_name(driver, registers)
                 for driver in ("R-001", "R-002", "CL-001", "CL-002", "CL-003")}
    for driver, name in published.items():
        assert isinstance(name, str), f"{driver} publishes a {type(name).__name__}"
        assert name.strip(), f"{driver} publishes an empty label"
        assert name != "0", f"{driver} publishes a numeric zero as its label"
    assert published["R-001"] == "GateB R-001"
    assert published["R-002"] == "GateB R-002"
    assert len(set(published.values())) == 5, "two drivers share a label"


def test_45_the_correction_touched_the_label_and_nothing_else() -> None:
    """THE MEASURE FIELDS, THE IDENTITY FIELDS AND THE ORDERING ARE UNTOUCHED.
    A label fix that moved a rho, a rank or the sort would be a different change
    wearing this one's authorisation."""
    fill = _procedure("FillRecord")
    # THE THREE IDENTITY FIELDS still come from where they came from.
    assert "record.PermanentId" in fill
    assert "DriverTypeOf(record.PermanentId)" in fill
    assert "DriverNameOf(record.PermanentId)" in fill
    # THE MEASURES ARE STILL THE RECORD'S OWN, not recomputed or re-signed.
    for measure in ("record.Rho", "record.AbsRho"):
        assert measure in fill, f"{measure} is no longer written from the record"
    assert "rank" in fill
    # AND NOTHING IN THE MODULE RE-RANKS OR RE-SIGNS while labelling.
    name = _procedure("DriverNameOf")
    for banned in ("Rho", "Rank", "Sort", "Abs", "Fingerprint", "Digest"):
        assert banned not in name, (
            f"the label lookup reaches into {banned}")
    # THE ONE COLUMN THAT MOVED, and it moved to the constant that owns it.
    assert "COL_RISK_REGISTER_RISK_NAME" in name
    assert "COL_COST_LINES_DESCRIPTION" in name
    assert "COL_RISK_REGISTER_DESCRIPTION" not in name, (
        "the risk label reads the optional description again")
    # NO NUMERIC COLUMN LITERAL: the ordinals stay the generated constants'.
    assert not re.search(r"column\s*=\s*\d+", name), (
        "a column ordinal is typed into the label lookup")


def test_46_the_contract_is_what_makes_the_choice_of_column_right() -> None:
    """NOT A PREFERENCE. Each register's label is the column its own contract
    declares required, and the two registers do not agree on which key that is."""
    import yaml
    contract = yaml.safe_load(
        (SPEC / "driver_contract.yaml").read_text(encoding="utf-8"))
    registers = contract["registers"]

    def column(register: str, key: str) -> dict:
        return next(c for c in registers[register]["columns"] if c["key"] == key)

    assert column("risk_register", "risk_name")["required"] is True
    assert column("risk_register", "risk_name")["type"] == "text"
    assert column("risk_register", "description")["required"] is False, (
        "the risk description became required; the label question reopens")
    assert column("cost_lines", "description")["required"] is True, (
        "the cost line description stopped being its required label")
    # A COST LINE HAS NO NAME COLUMN TO MOVE TO, which is why only one side
    # of the lookup changed.
    assert not any(c["key"] == "name" or c["key"] == "cost_line_name"
                   for c in registers["cost_lines"]["columns"])


# ---------------------------------------------------------------------------
# THE MUTATIONS
# ---------------------------------------------------------------------------
# EACH IS A WAY THE LABEL CORRECTION COULD LOOK DONE AND NOT BE. They are
# applied to a COPY of the module source in memory and the rules above are
# re-run over it; nothing on disk changes. Four of them are the pre-fix defect
# in a different costume, and three are the ways a fix could overreach.
def _label_rules(source: str) -> None:
    """Everything this correction claims, checked over an arbitrary copy."""
    name = re.search(r"^Private Function DriverNameOf.*?^End Function",
                     source, re.S | re.M)
    assert name, "the label lookup is gone"
    name = name.group(0)
    # THE RISK READS ITS REQUIRED NAME, THE COST LINE ITS REQUIRED DESCRIPTION.
    assert "COL_RISK_REGISTER_RISK_NAME" in name, (
        "the risk label does not read the risk name")
    assert "COL_RISK_REGISTER_DESCRIPTION" not in name, (
        "the risk label reads the optional description")
    assert "COL_COST_LINES_DESCRIPTION" in name, (
        "the cost-line label stopped reading its description")
    # NEITHER IS THE IDENTIFIER.
    for identifier in ("COL_RISK_REGISTER_RISK_ID", "COL_COST_LINES_COST_LINE_ID"):
        assert identifier not in name, f"the label publishes {identifier}"
    # NO NUMERIC ORDINAL, and no invented value for a name that is not there.
    assert not re.search(r"column\s*=\s*\d+", name), (
        "a column ordinal is typed instead of named")
    for invented in ('= "0"', "= \"unnamed\"", "= permanentId", "IIf("):
        assert invented not in name, (
            f"a missing label is coerced to an invented value: {invented}")
    # AND THE LOOKUP STAYS A LOOKUP.
    for reach in ("Rho", "Rank", "Sort", "AbsRho", "Fingerprint", "Digest"):
        assert reach not in name, f"the label lookup reaches into {reach}"
    # THE MEASURES AND THE ORDER ARE UNTOUCHED.
    fill = re.search(r"^Private Sub FillRecord.*?^End Sub", source, re.S | re.M)
    assert fill, "the record writer is gone"
    fill = fill.group(0)
    assert "block(row, SIM_SENSITIVITY_OFFSET_RHO + 1) = record.Rho" in fill, (
        "the signed rho is no longer written from the record")
    assert "block(row, SIM_SENSITIVITY_OFFSET_ABS_RHO + 1) = record.AbsRho" in fill
    assert "block(row, SIM_SENSITIVITY_OFFSET_RANK + 1) = rank" in fill, (
        "the rank is no longer the one it was handed")
    assert "DirectionOf(record.Rho)" in fill, "the direction stopped following the sign"


@pytest.mark.parametrize("name,mutate", [
    # THE DEFECT ITSELF, RESTORED.
    ("the risk label reads the optional description again",
     lambda src: src.replace("    column = COL_RISK_REGISTER_RISK_NAME",
                             "    column = COL_RISK_REGISTER_DESCRIPTION", 1)),
    # THE RIGHT COLUMN BY THE WRONG ROUTE. A literal 2 is correct today and
    # silently wrong the moment the register gains a column.
    ("the risk name column is hard-coded",
     lambda src: src.replace("    column = COL_RISK_REGISTER_RISK_NAME",
                             "    column = 2", 1)),
    # THE IDENTIFIER PUBLISHED AS THE LABEL - never empty, and never a name.
    ("the driver id is published as the driver name",
     lambda src: src.replace("    column = COL_RISK_REGISTER_RISK_NAME",
                             "    column = COL_RISK_REGISTER_RISK_ID", 1)),
    # THE COST-LINE SIDE MOVED TOO, which no contract asks for.
    ("the cost-line label stops reading its description",
     lambda src: src.replace("        column = COL_COST_LINES_DESCRIPTION",
                             "        column = COL_COST_LINES_CATEGORY", 1)),
    # THE SYMPTOM PAPERED OVER INSTEAD OF THE OWNERSHIP FIXED.
    ("a missing label is coerced to a value",
     lambda src: src.replace(
         "    DriverNameOf = modWorkbook.TextOf(modWorkbook.CellIn(table, row, column))",
         "    DriverNameOf = modWorkbook.TextOf(modWorkbook.CellIn(table, row, column))\n"
         "    If Len(DriverNameOf) = 0 Then DriverNameOf = \"0\"", 1)),
    # THE SIGN CHANGED WHILE THE LABEL WAS BEING FIXED.
    ("the signed rho is published as a magnitude",
     lambda src: src.replace(
         "        block(row, SIM_SENSITIVITY_OFFSET_RHO + 1) = record.Rho",
         "        block(row, SIM_SENSITIVITY_OFFSET_RHO + 1) = record.AbsRho", 1)),
    # THE RANK REWRITTEN WHILE THE LABEL WAS BEING FIXED.
    ("the rank stops being the one the ranking assigned",
     lambda src: src.replace(
         "        block(row, SIM_SENSITIVITY_OFFSET_RANK + 1) = rank",
         "        block(row, SIM_SENSITIVITY_OFFSET_RANK + 1) = row", 1)),
])
def test_47_each_way_of_getting_the_label_wrong_is_refused(name: str, mutate) -> None:
    source = POST_BAS.read_text(encoding="utf-8")
    mutated = mutate(source)
    assert mutated != source, f"the mutation '{name}' changed nothing"
    with pytest.raises(AssertionError):
        _label_rules(mutated)


def test_48_the_label_rules_pass_on_the_real_module() -> None:
    """SO THE SEVEN REFUSALS ABOVE ARE REFUSALS OF THE MUTATION, not of the
    fixture."""
    _label_rules(POST_BAS.read_text(encoding="utf-8"))


# ===========================================================================
# J. ZERO VARIANCE IS UNDEFINED, NOT ZERO - P8-3 CLOSURE
# ===========================================================================
# THE PATH, END TO END. A driver whose contribution never varies has no monotone
# association to find, so the kernel gives it no rho. It is still PUBLISHED -
# Publish writes the eligible drivers in ranked order and then appends every
# record whose Status is not SIM_SENSITIVITY_DEFINED - and FillRecord's Else
# branch deliberately writes vbNullString to rho, |rho|, rank and direction,
# with only the status label carrying "n/a - no variance".
#
# WHERE IT WENT WRONG. The Sensitivity sheet reads each field as a LOOKUP into
# the active bank, and the innermost term was a BARE REFERENCE. Excel reads an
# empty reference back as ZERO, so all four deliberately-blank fields arrived on
# the sheet as a measured 0 - and the chart bridge's own `=""` guard cannot
# catch a number, so a driver with no measurable association reached the tornado
# as a zero-length bar.
#
# THE CONTRACT SAYS THIS IN AS MANY WORDS: "Zero variance is UNDEFINED, not
# zero. Reporting rho = 0 for a constant column asserts 'no monotone association
# was found', which is a measurement. No measurement was possible."
def _sensitivity_cell(column_key: str, row_offset: int = 0) -> str:
    shell = _shell()
    column = next(c for c in shell["columns"] if c["key"] == column_key)
    return _built_sheet()[
        f"{column['column']}{int(shell['first_row']) + row_offset}"].value


def test_50_the_contract_forbids_a_zero_variance_driver_from_reporting_zero() -> None:
    """THE RULE THIS EXISTS TO KEEP. Read from the contract, not restated."""
    zero_variance = _raw_contract()["sensitivity"]["zero_variance"]
    assert zero_variance["rho_reported"] is False
    assert zero_variance["reported_as_zero_rho"] is False
    assert zero_variance["excluded_from_ranking"] is True
    assert zero_variance["excluded_from_tornado_input"] is True
    assert zero_variance["retained_diagnostically"] is True, (
        "a zero-variance driver stopped being reported at all; it is excluded "
        "from the RANKING, not from the table")
    assert zero_variance["status_label"] == "n/a - no variance"


def test_51_a_zero_variance_record_is_published_with_no_measure() -> None:
    """SO THE BLANKS ARE REAL AND THEY REACH THE SHEET. This is what makes the
    presentation guard necessary rather than defensive."""
    publish = _procedure("Publish")
    # THE ELIGIBLE ONES FIRST, THEN EVERYTHING ELSE - so a zero-variance driver
    # occupies a published row, and one inside the first rows whenever fewer
    # than that many drivers are eligible.
    assert "For position = 0 To eligibleCount - 1" in publish
    assert "If results(index).Status <> SIM_SENSITIVITY_DEFINED Then" in publish, (
        "the ineligible records are no longer appended")
    assert "SIM_SENSITIVITY_STAMP_ROW_RECORD_COUNT).Value2 = driverCount" in publish, (
        "the persisted count no longer covers the diagnostic rows")
    fill = _procedure("FillRecord")
    undefined = fill[fill.index("    Else"):]
    for measure in ("RHO", "ABS_RHO", "RANK", "DIRECTION"):
        assert f"SIM_SENSITIVITY_OFFSET_{measure} + 1) = vbNullString" in undefined, (
            f"{measure} is no longer left blank for a zero-variance driver")
    assert "SENSITIVITY_NO_VARIANCE_LABEL" in undefined


def _record_reference(column_key: str, row_offset: int = 0) -> str:
    """The exact bank-picking expression the sheet must guard, DERIVED from the
    contract coordinates rather than read back out of the formula.

    Deriving it is the point: a control that lifted the expression out of the
    formula it is checking would agree with whatever it found there."""
    records = _raw_contract()["sim_data"]["sensitivity_records"]
    sheet = _raw_contract()["sim_data"]["sheet"]
    columns = {c["key"]: int(c["offset"]) if "offset" in c else None
               for c in _shell()["columns"]}
    offset = columns[column_key]
    assert offset is not None, column_key

    def shift(letter: str, by: int) -> str:
        value = 0
        for char in letter:
            value = value * 26 + (ord(char) - ord("A") + 1)
        value += by
        out = ""
        while value:
            value, remainder = divmod(value - 1, 26)
            out = chr(ord("A") + remainder) + out
        return out

    row = int(records["first_record_row"]) + row_offset
    banks = records["banks"]
    active = f'{sheet}!$D$30'
    a = f'{sheet}!${shift(banks["A"]["first_column"], offset)}${row}'
    b = f'{sheet}!${shift(banks["B"]["first_column"], offset)}${row}'
    return f'IF({active}="A",{a},{b})'


@pytest.mark.parametrize("column_key", ["rho", "abs_rho", "rank", "direction",
                                        "driver_name", "driver_id", "driver_type",
                                        "status"])
def test_52_every_sensitivity_field_guards_its_empty_reference(column_key: str) -> None:
    """EXCEL READS AN EMPTY REFERENCE BACK AS ZERO, so every field guards it -
    not only the four a zero-variance row leaves blank. A field that is blank
    for any reason must arrive blank.

    THE ASSERTION IS ON THE RECORD REFERENCE ITSELF. The formula's outer bounds
    already contain `=""` three times over, so testing the formula for that
    string proves nothing - which is exactly what the first version of this
    control did, and it passed against the unguarded build."""
    formula = _sensitivity_cell(column_key)
    picked = _record_reference(column_key)
    assert picked in formula, (
        f"{column_key} does not read the record the contract places it at: "
        f"{picked}")
    guarded = f'IF({picked}="","",{picked})'
    assert guarded in formula, (
        f"{column_key} reads its record UNGUARDED; Excel returns 0 for an empty "
        f"reference and a deliberately blank field becomes a measured zero.\n"
        f"  wanted: {guarded}\n  formula: {formula}")


def test_53_the_guard_did_not_swallow_a_measured_zero() -> None:
    """A DRIVER WHOSE RHO REALLY IS 0 STILL SHOWS 0. `0=""` is FALSE in Excel,
    so the guard tests emptiness and never a value - which is the whole
    difference between "no measurement" and "a measurement of zero"."""
    formula = _sensitivity_cell("rho")
    picked = _record_reference("rho")
    assert f'IF({picked}="","",{picked})' in formula
    for numeric in (f'IF({picked}=0', 'ISNUMBER(', 'ISBLANK(', 'N(', 'ABS(',
                    'IFERROR('):
        assert numeric not in formula, (
            f"the guard tests a value rather than emptiness: {numeric}")


def test_54_the_four_blank_bounds_are_still_blank_and_still_bounded() -> None:
    """THE THREE CONDITIONS THAT WERE ALREADY THERE ARE UNTOUCHED. The guard is
    a fourth; it did not replace the publication, count or activity bounds."""
    formula = _sensitivity_cell("rho")
    shell = _shell()
    records = _raw_contract()["sim_data"]["sensitivity_records"]
    count_row = next(f["row"] for f in records["stamp"]["fields"]
                     if f["key"] == "record_count")
    published_row = next(f["row"] for f in records["stamp"]["fields"]
                         if f["key"] == "published")
    assert '<>"PUBLISHED"' in formula, "the publication bound was lost"
    assert f"${count_row}" in formula, "the record-count bound was lost"
    assert f"${published_row}" in formula, "the published-stamp bound was lost"
    assert formula.count('IF(') >= 4, "the guard replaced a bound instead of adding one"
    # AND THE WINDOW IS STILL THE MANIFEST'S.
    assert int(shell["row_window"]) >= 1


def test_55_the_tornado_bridge_now_receives_a_blank_and_draws_nothing() -> None:
    """THE END OF THE CHAIN. The bridge refuses to plot where its source is
    blank; it could never refuse a number. With the sheet blank, the refusal
    fires - and this is checked at the bridge WITHOUT changing it."""
    import openpyxl
    from pccm_builder.spec_loader import load_spec
    charts = load_spec(SPEC / "workbook.yaml").phase6_shell["charts"]
    drivers = charts["bridge"]["drivers"]
    workbook = openpyxl.load_workbook(PCCM_ROOT / "build" / "PCCM_stageA.xlsx")
    sheet = workbook[charts["bridge_sheet"]]
    rho = next(c for c in drivers["columns"] if c["key"] == "rho")
    formula = sheet[f"{rho['column']}{int(drivers['first_row'])}"].value
    assert 'NA()' in formula, "the bridge stopped refusing an absent driver"
    assert '=""' in formula, "the bridge no longer tests its source for emptiness"
    # THE BRIDGE IS UNCHANGED - the correction is upstream of it.
    assert formula.startswith(f"=IF({charts['sensitivity_sheet']}!")


# ---------------------------------------------------------------------------
# THE MUTATIONS - each is a way the guard could look present and not be
# ---------------------------------------------------------------------------
def _rendered(mutate=None) -> dict[str, str]:
    """The Sensitivity row-1 formulas the builder produces, optionally from a
    mutated copy of its source.

    ONLY THE RENDERER IS RECOMPILED, over the real module's own globals. Exec-ing
    the whole module would drag in its imports and dataclasses and prove nothing
    extra; rebinding the one function under test against everything else as it
    really is keeps the mutation the only difference. Nothing on disk changes
    and no workbook is written."""
    from pccm_builder import workbook_builder as real

    source = Path(real.__file__).read_text(encoding="utf-8")
    if mutate is not None:
        mutated = mutate(source)
        assert mutated != source, "the mutation changed nothing"
        source = mutated
    body = re.search(r"^def _render_sensitivity_shell\(.*?(?=\n\ndef )",
                     source, re.S | re.M)
    assert body, "the sensitivity renderer is gone"
    namespace = dict(real.__dict__)
    exec(compile(body.group(0), real.__file__, "exec"), namespace)

    written: dict[str, str] = {}

    class _Cell:
        def __init__(self, key):
            self._key = key

        @property
        def value(self):
            return written.get(self._key)

        @value.setter
        def value(self, new):
            written[self._key] = new

    class _Sheet:
        def __getitem__(self, address):
            return _Cell(address)

    from pccm_builder.spec_loader import load_spec
    shell = _shell()
    styles = real.StyleBook(load_spec(SPEC / "workbook.yaml").presentation)
    namespace["_render_sensitivity_shell"](_Sheet(), shell, _raw_contract(), styles)
    first = int(shell["first_row"])
    return {c["key"]: written.get(f"{c['column']}{first}", "")
            for c in shell["columns"]}


def _guard_rules(rendered: dict[str, str]) -> None:
    """Every field guards its record reference, and the bounds are still there."""
    for key, formula in rendered.items():
        picked = _record_reference(key)
        assert picked in formula, f"{key} does not read its contracted record"
        assert f'IF({picked}="","",{picked})' in formula, (
            f"{key} reads its record unguarded")
        assert '<>"PUBLISHED"' in formula, f"{key} lost the publication bound"


@pytest.mark.parametrize("name,mutate", [
    # THE DEFECT ITSELF, RESTORED.
    ("the record reference is read unguarded",
     lambda src: src.replace('f\'IF({picked}="","",{picked}))))\'',
                             'f\'{picked})))\'', 1)),
    # THE GUARD TESTING THE WRONG THING. `0=0` is TRUE, so a genuinely measured
    # zero would be blanked - the opposite error and just as wrong.
    ("the guard blanks a measured zero",
     lambda src: src.replace('f\'IF({picked}="","",{picked}))))\'',
                             'f\'IF({picked}=0,"",{picked}))))\'', 1)),
    # THE GUARD APPLIED TO SOMETHING THAT IS NEVER EMPTY, which is the same as
    # not guarding: the active-bank selector always holds a value.
    ("the guard tests the bank selector instead of the record",
     lambda src: src.replace('f\'IF({picked}="","",{picked}))))\'',
                             'f\'IF({active}="","",{picked}))))\'', 1)),
    # A PUBLICATION BOUND TRADED FOR THE GUARD rather than added to it.
    ("the guard replaced the publication bound",
     lambda src: src.replace(
         'f\'IF(IF({active}="A",{stamp_cell("A", published_row)},\'\n'
         '                f\'{stamp_cell("B", published_row)})<>"PUBLISHED","",\'',
         '', 1)),
])
def test_56_each_way_of_losing_the_guard_is_refused(name: str, mutate) -> None:
    with pytest.raises(AssertionError):
        _guard_rules(_rendered(mutate))


def test_57_the_guard_rules_pass_on_the_real_builder() -> None:
    """SO THE FOUR REFUSALS ABOVE ARE REFUSALS OF THE MUTATION, not of the
    fixture - and the renderer really does produce what the built workbook holds."""
    rendered = _rendered()
    _guard_rules(rendered)
    for key, formula in rendered.items():
        assert formula == _sensitivity_cell(key), (
            f"{key}: the renderer and the built workbook disagree")


# ===========================================================================
# K. THE TORNADO INPUT IS THE RANKED POPULATION - P8-3 FINAL CLOSURE
# ===========================================================================
# THE CONTRACT CLAUSE, READ LITERALLY. `zero_variance` carries five sibling
# statements, and `excluded_from_tornado_input: true` sits beside
# `rho_reported: false` and `reported_as_zero_rho: false`. Read as "its bar is
# blank" it would restate those two and mean nothing of its own. The only
# reading under which it says anything is that the driver is not part of the
# tornado's input AT ALL - neither its value nor its identity. `ranking` says
# the same from the other side: population is "every eligible non_zero_variance
# driver".
#
# WHY THE BRIDGE VIOLATED IT. Publish writes the eligible drivers in ranked
# order and then APPENDS the diagnostic ones, so "the first N rows of the sheet"
# equals "the first N of the ranked population" only while at least N drivers
# are eligible. Below that they diverge, and a zero-variance driver entered the
# category window - after 94c6b37 with no bar, but still as a category.
#
# THESE TESTS EVALUATE THE FORMULAS. A tiny Excel-subset evaluator resolves the
# real generated bridge and Sensitivity formulas against a simulated _SimData,
# so the assertions are about what the workbook COMPUTES, not what its text
# looks like.
ZERO_VARIANCE_STATUS = "n/a - no variance"


def _sensitivity_columns() -> dict[str, str]:
    return {str(c["key"]): str(c["column"]) for c in _shell()["columns"]}


class _Book:
    """A tiny evaluator for the exact formula shapes these two layers emit.

    IT UNDERSTANDS NOTHING ELSE. `IF`, `NA()`, `=`, `<>`, a cell reference and a
    string literal - which is every construct in the generated bridge and
    Sensitivity formulas - and it raises on anything it does not recognise, so a
    formula that grew a new shape cannot be silently mis-evaluated.
    """

    NA = object()

    def __init__(self, cells: dict[str, object]):
        self.cells = cells

    def value(self, ref: str):
        key = ref.replace("$", "")
        if key not in self.cells:
            # AN UNWRITTEN CELL IS EMPTY, and Excel reads an empty reference
            # back as 0. Modelling that is the whole reason this exists.
            return 0
        held = self.cells[key]
        return "" if held is None else held

    def evaluate(self, formula: str):
        assert formula.startswith("="), formula
        result = self._expr(formula[1:])
        return result

    def _split(self, text: str) -> list[str]:
        parts, depth, quoted, token = [], 0, False, ""
        for char in text:
            if char == '"':
                quoted = not quoted
            if not quoted:
                if char == "(":
                    depth += 1
                elif char == ")":
                    depth -= 1
                elif char == "," and depth == 0:
                    parts.append(token)
                    token = ""
                    continue
            token += char
        parts.append(token)
        return parts

    def _expr(self, text: str):
        text = text.strip()
        if text == "NA()":
            return self.NA
        if text.startswith("IF(") and text.endswith(")"):
            condition, when_true, when_false = self._split(text[3:-1])
            return self._expr(when_true if self._condition(condition)
                              else when_false)
        if text.startswith('"') and text.endswith('"'):
            return text[1:-1]
        if re.fullmatch(r"-?\d+(\.\d+)?", text):
            return float(text) if "." in text else int(text)
        if re.fullmatch(r"[A-Za-z_]+!\$?[A-Z]{1,3}\$?\d+", text):
            return self.value(text.split("!", 1)[1])
        raise AssertionError(f"the evaluator does not understand {text!r}")

    def _condition(self, text: str) -> bool:
        for operator in ("<>", ">", "="):
            index, depth, quoted = -1, 0, False
            for position, char in enumerate(text):
                if char == '"':
                    quoted = not quoted
                if quoted:
                    continue
                if char == "(":
                    depth += 1
                elif char == ")":
                    depth -= 1
                elif depth == 0 and text.startswith(operator, position):
                    index = position
                    break
            if index < 0:
                continue
            left = self._expr(text[:index])
            right = self._expr(text[index + len(operator):])
            if operator == "=":
                return left == right
            if operator == "<>":
                return left != right
            return (left or 0) > (right or 0)
        raise AssertionError(f"the evaluator does not understand {text!r}")


def _zero_variance_book(drivers: list[dict]) -> _Book:
    """A published _SimData bank A holding these records, in Publish's own order.

    RANKED FIRST, DIAGNOSTIC APPENDED - the order the module writes, reproduced
    rather than assumed: eligible records carry a rank, the rest carry blanks
    for every measure and the zero-variance status label.
    """
    records = _raw_contract()["sim_data"]["sensitivity_records"]
    sheet = _raw_contract()["sim_data"]["sheet"]
    first_record_row = int(records["first_record_row"])
    stamp = {f["key"]: f["row"] for f in records["stamp"]["fields"]}
    stamp_col = records["stamp"]["bank_value_columns"]["A"]
    offsets = {str(c["key"]): int(c["offset"]) for c in _shell()["columns"]}

    def shift(letter: str, by: int) -> str:
        value = 0
        for char in letter:
            value = value * 26 + (ord(char) - ord("A") + 1)
        value += by
        out = ""
        while value:
            value, remainder = divmod(value - 1, 26)
            out = chr(ord("A") + remainder) + out
        return out

    bank_first = records["banks"]["A"]["first_column"]
    cells: dict[str, object] = {
        "D30": "A",
        f"{stamp_col}{stamp['published']}": "PUBLISHED",
        f"{stamp_col}{stamp['record_count']}": len(drivers),
    }
    eligible = [d for d in drivers if d["status"] != ZERO_VARIANCE_STATUS]
    diagnostic = [d for d in drivers if d["status"] == ZERO_VARIANCE_STATUS]
    for position, driver in enumerate(eligible + diagnostic):
        row = first_record_row + position
        ranked = driver["status"] != ZERO_VARIANCE_STATUS
        fields = {
            "driver_id": driver["id"],
            "driver_type": driver.get("type", "Cost"),
            "driver_name": driver["name"],
            "rho": driver["rho"] if ranked else None,
            "abs_rho": abs(driver["rho"]) if ranked else None,
            "rank": position + 1 if ranked else None,
            "direction": ("-" if ranked and driver["rho"] < 0 else
                          "+" if ranked else None),
            "status": driver["status"],
        }
        for key, held in fields.items():
            cells[f"{shift(bank_first, offsets[key])}{row}"] = held
    return _Book(cells), sheet


def _resolve(book: _Book, formula: str):
    return book.evaluate(formula)


# THE FIXTURE §4 REQUIRES: several DEFINED drivers, one of them with a genuine
# measured rho of exactly 0, one zero-variance driver, and fewer than ten
# DEFINED drivers in total.
ZERO_VARIANCE_FIXTURE = [
    {"id": "CL-001", "name": "Strong positive", "rho": 0.81, "status": "ranked"},
    {"id": "R-001", "name": "Strong negative", "rho": -0.64, "status": "ranked"},
    {"id": "CL-002", "name": "Weak positive", "rho": 0.12, "status": "ranked"},
    # A MEASURED ZERO. The driver varied and no monotone association was found -
    # which is a RESULT, and entirely different from having none to look for.
    {"id": "CL-003", "name": "Measured zero", "rho": 0.0, "status": "ranked"},
    # AND THE ONE THAT MUST NOT REACH THE CHART.
    {"id": "R-002", "name": "Constant impact", "rho": None,
     "status": ZERO_VARIANCE_STATUS},
]


def _tornado_rows() -> list[dict]:
    """What the tornado bridge actually resolves to, row by row, over the
    fixture - using the REAL generated formulas from the built workbook."""
    import openpyxl
    from pccm_builder.spec_loader import load_spec

    book, _ = _zero_variance_book(ZERO_VARIANCE_FIXTURE)
    charts = load_spec(SPEC / "workbook.yaml").phase6_shell["charts"]
    drivers = charts["bridge"]["drivers"]
    workbook = openpyxl.load_workbook(PCCM_ROOT / "build" / "PCCM_stageA.xlsx")
    results = workbook[charts["bridge_sheet"]]
    sensitivity = workbook[charts["sensitivity_sheet"]]
    columns = {str(c["key"]): str(c["column"]) for c in drivers["columns"]}
    first = int(drivers["first_row"])

    def sheet_value(address: str):
        return _resolve(book, sensitivity[address].value)

    rows = []
    for index in range(int(drivers["top_n"])):
        row = {}
        for key, column in columns.items():
            formula = results[f"{column}{first + index}"].value
            # THE BRIDGE READS THE SENSITIVITY SHEET, which is itself a formula.
            resolved = formula
            for match in sorted(set(re.findall(
                    r"Sensitivity!\$([A-Z]{1,3})\$(\d+)", formula)),
                    key=lambda m: -(len(m[0]) + len(m[1]))):
                value = sheet_value(f"{match[0]}{match[1]}")
                literal = ('""' if value == "" else
                           f'"{value}"' if isinstance(value, str) else str(value))
                resolved = resolved.replace(
                    f"Sensitivity!${match[0]}${match[1]}", literal)
            row[key] = _Book({}).evaluate(resolved)
        rows.append(row)
    return rows


def test_60_the_contract_excludes_a_zero_variance_driver_from_the_tornado() -> None:
    """THE LITERAL READING. `excluded_from_tornado_input` is a SEPARATE clause
    from the two beside it, so it cannot mean what they already say."""
    zero_variance = _raw_contract()["sensitivity"]["zero_variance"]
    assert zero_variance["excluded_from_tornado_input"] is True
    # THE TWO CLAUSES IT WOULD OTHERWISE RESTATE.
    assert zero_variance["rho_reported"] is False
    assert zero_variance["reported_as_zero_rho"] is False
    # AND THE RANKING SAYS IT FROM THE OTHER SIDE.
    ranking = _raw_contract()["sensitivity"]["ranking"]
    assert ranking["population"] == "every eligible non_zero_variance driver"
    assert ranking["top_n_truncation"] is False, (
        "Phase 7 stopped producing the whole ranked population; Top-N is the "
        "chart's choice to make and it needs everything to choose from")


def test_61_the_ranked_drivers_all_appear_in_published_order() -> None:
    """POSITIVE, NEGATIVE, AND A GENUINELY MEASURED ZERO."""
    rows = _tornado_rows()
    ranked = [d for d in ZERO_VARIANCE_FIXTURE
              if d["status"] != ZERO_VARIANCE_STATUS]
    for index, driver in enumerate(ranked):
        assert rows[index]["driver_name"] == driver["name"], (
            f"row {index + 1} is not the published driver of that rank")
        assert rows[index]["rho"] == driver["rho"], (
            f"{driver['name']}: rho {rows[index]['rho']} != {driver['rho']}")
    # THE SIGN SURVIVES, and the order is the sheet's, not a re-sort.
    assert rows[0]["rho"] > 0 and rows[1]["rho"] < 0
    assert [r["driver_name"] for r in rows[:len(ranked)]] == [
        d["name"] for d in ranked]


def test_62_a_measured_zero_is_not_an_undefined_one() -> None:
    """THE DISTINCTION THE WHOLE CLAUSE RESTS ON. A driver that varied and
    showed no monotone association HAS a result; one with nothing to correlate
    has none. The first is plotted at zero; the second is absent."""
    rows = _tornado_rows()
    measured = next(i for i, d in enumerate(ZERO_VARIANCE_FIXTURE)
                    if d["name"] == "Measured zero")
    assert rows[measured]["rho"] == 0.0, "the measured zero was excluded"
    assert rows[measured]["driver_name"] == "Measured zero"
    assert rows[measured]["rho"] is not _Book.NA


def test_63_the_zero_variance_driver_is_not_a_category_and_not_a_bar() -> None:
    """NEITHER ITS VALUE NOR ITS IDENTITY reaches the chart."""
    rows = _tornado_rows()
    absent = {"Constant impact"}
    for row in rows:
        assert row["driver_name"] not in absent, (
            "the zero-variance driver is still a tornado category")
    ranked = [d for d in ZERO_VARIANCE_FIXTURE
              if d["status"] != ZERO_VARIANCE_STATUS]
    for row in rows[len(ranked):]:
        assert row["driver_name"] is _Book.NA, (
            "a row beyond the ranked population carries a category")
        assert row["rho"] is _Book.NA, (
            "a row beyond the ranked population carries a value")


def test_64_fewer_eligible_drivers_yield_fewer_categories_and_no_filler() -> None:
    rows = _tornado_rows()
    ranked = [d for d in ZERO_VARIANCE_FIXTURE
              if d["status"] != ZERO_VARIANCE_STATUS]
    assert len(ranked) < 10, "the fixture no longer exercises the short case"
    plotted = [r for r in rows if r["rho"] is not _Book.NA]
    assert len(plotted) == len(ranked), (
        f"{len(plotted)} categories for {len(ranked)} eligible drivers")
    # NO FILLER OF ANY KIND, and in particular not a zero or a blank.
    for row in rows[len(ranked):]:
        assert row["rho"] is _Book.NA and row["driver_name"] is _Book.NA
        assert row["rho"] != 0 and row["driver_name"] != ""
    # AND THE WINDOW IS STILL AT MOST TEN.
    from pccm_builder.spec_loader import load_spec
    drivers = load_spec(SPEC / "workbook.yaml").phase6_shell["charts"]["bridge"]["drivers"]
    assert int(drivers["top_n"]) == 10
    assert len(rows) == 10


def test_65_the_diagnostic_row_is_still_on_the_sensitivity_sheet() -> None:
    """RETAINED DIAGNOSTICALLY. Excluding it from the CHART is not removing it
    from the record: the sheet still shows the driver, its type, its name and
    the status that says why it has no rho."""
    assert _raw_contract()["sensitivity"]["zero_variance"][
        "retained_diagnostically"] is True
    import openpyxl
    from pccm_builder.spec_loader import load_spec
    book, _ = _zero_variance_book(ZERO_VARIANCE_FIXTURE)
    workbook = openpyxl.load_workbook(PCCM_ROOT / "build" / "PCCM_stageA.xlsx")
    sheet = workbook[load_spec(SPEC / "workbook.yaml").phase6_shell[
        "sensitivity"]["sheet"]]
    columns = _sensitivity_columns()
    first = int(_shell()["first_row"])
    # THE DIAGNOSTIC ROW IS THE LAST PUBLISHED ONE.
    row = first + len(ZERO_VARIANCE_FIXTURE) - 1

    def cell(key: str):
        return _resolve(book, sheet[f"{columns[key]}{row}"].value)

    assert cell("driver_id") == "R-002", cell("driver_id")
    assert cell("driver_name") == "Constant impact"
    assert cell("status") == ZERO_VARIANCE_STATUS
    # AND ITS MEASURES ARE BLANK, NOT ZERO - the 94c6b37 guard, still holding.
    for measure in ("rho", "abs_rho", "rank", "direction"):
        assert cell(measure) == "", (
            f"{measure} reads {cell(measure)!r}; a zero-variance driver reports "
            "no measurement, not a measurement of zero")


def test_66_the_bridge_gates_on_rank_and_not_on_a_blank_rho() -> None:
    """THE AUTHORITATIVE FIELD. rho being absent is a CONSEQUENCE of exclusion
    from the ranking; rank is where the publication writes the exclusion down."""
    import openpyxl
    from pccm_builder.spec_loader import load_spec
    charts = load_spec(SPEC / "workbook.yaml").phase6_shell["charts"]
    drivers = charts["bridge"]["drivers"]
    columns = _sensitivity_columns()
    workbook = openpyxl.load_workbook(PCCM_ROOT / "build" / "PCCM_stageA.xlsx")
    results = workbook[charts["bridge_sheet"]]
    first = int(drivers["first_row"])
    sheet = charts["sensitivity_sheet"]
    for index in range(int(drivers["top_n"])):
        gate = f'IF({sheet}!${columns["rank"]}${int(_shell()["first_row"]) + index}="",NA(),'
        for column in drivers["columns"]:
            formula = results[f"{column['column']}{first + index}"].value
            assert formula.startswith("=" + gate), (
                f"{column['key']} row {index + 1} is not gated on rank: {formula}")
    # AND THE ELIGIBILITY FIELD IS NOT ITSELF PLOTTED.
    assert "rank" not in {str(c["source"]) for c in drivers["columns"]}


def test_67_the_publication_keeps_rank_and_status_in_agreement() -> None:
    """SO GATING ON RANK IS GATING ON THE SAME POPULATION THE STATUS NAMES. The
    two are written in the same two branches and cannot drift apart."""
    fill = _procedure("FillRecord")
    defined = fill[fill.index("If record.Status = SIM_SENSITIVITY_DEFINED Then"):
                   fill.index("    Else")]
    undefined = fill[fill.index("    Else"):]
    assert "SIM_SENSITIVITY_OFFSET_RANK + 1) = rank" in defined
    assert "SENSITIVITY_RANKED_LABEL" in defined
    assert "SIM_SENSITIVITY_OFFSET_RANK + 1) = vbNullString" in undefined
    assert "SENSITIVITY_NO_VARIANCE_LABEL" in undefined
    # AND THE ZERO-VARIANCE LABEL IS THE CONTRACT'S.
    module = POST_BAS.read_text(encoding="utf-8")
    assert f'SENSITIVITY_NO_VARIANCE_LABEL As String = "{ZERO_VARIANCE_STATUS}"' \
        in module
    assert _raw_contract()["sensitivity"]["zero_variance"]["status_label"] == \
        ZERO_VARIANCE_STATUS


# ---------------------------------------------------------------------------
# THE MUTATIONS
# ---------------------------------------------------------------------------
def _bridge_rows(mutate=None) -> list[tuple[int, str, str]]:
    """The tornado bridge formulas the builder emits, optionally from a mutated
    copy of its source. Only the two functions under test are recompiled, over
    the real module's globals."""
    from pccm_builder import workbook_builder as real
    from pccm_builder.spec_loader import load_spec

    source = Path(real.__file__).read_text(encoding="utf-8")
    if mutate is not None:
        mutated = mutate(source)
        assert mutated != source, "the mutation changed nothing"
        source = mutated
    body = re.search(r"^TORNADO_ELIGIBILITY_FIELD = .*?^def _chart_bridge_drivers"
                     r"\(.*?(?=\n\n(?:#|[A-Z_]+ =|def ))", source, re.S | re.M)
    assert body, "the tornado bridge builder is gone"
    namespace = dict(real.__dict__)
    exec(compile(body.group(0), real.__file__, "exec"), namespace)
    charts = load_spec(SPEC / "workbook.yaml").phase6_shell["charts"]
    return namespace["_chart_bridge_drivers"](
        charts["bridge"]["drivers"], _shell())


def _tornado_input_rules(rows: list[tuple[int, str, str]]) -> None:
    """The tornado plots the ranked population, in the sheet's own order."""
    from pccm_builder.spec_loader import load_spec
    charts = load_spec(SPEC / "workbook.yaml").phase6_shell["charts"]
    drivers = charts["bridge"]["drivers"]
    sheet = charts["sensitivity_sheet"]
    columns = _sensitivity_columns()
    first_source = int(_shell()["first_row"])
    top_n = int(drivers["top_n"])
    assert top_n <= 10, f"the tornado window grew to {top_n}"
    assert len(rows) == top_n * len(drivers["columns"]), (
        f"{len(rows)} cells for {top_n} rows of {len(drivers['columns'])} fields")
    by_row: dict[int, list[str]] = {}
    for row, _column, formula in rows:
        by_row.setdefault(row, []).append(formula)
    for index, row in enumerate(sorted(by_row)):
        gate = f'{sheet}!${columns["rank"]}${first_source + index}'
        for formula in by_row[row]:
            # EVERY FIELD IS GATED, the category as much as the value.
            assert formula.startswith(f'=IF({gate}="",NA(),'), (
                f"row {index + 1} is not gated on the ranked position: {formula}")
            # AND THE ROW IT READS IS THE POSITIONAL ONE - no re-sort, no
            # compaction, no ABS ordering.
            assert f"${first_source + index}" in formula, (
                f"row {index + 1} does not read source row {first_source + index}")
            for banned in ("LARGE(", "SMALL(", "RANK(", "ABS(", "SORT(",
                           "INDEX(", "MATCH(", "AGGREGATE("):
                assert banned not in formula, (
                    f"the bridge re-derives the order: {banned}")
            # NO FILLER. An excluded row is absent, never a zero or a blank.
            assert ',NA(),' in formula and ',0)' not in formula, formula
    # THE ELIGIBILITY FIELD IS NOT PLOTTED.
    assert "rank" not in {str(c["source"]) for c in drivers["columns"]}


@pytest.mark.parametrize("name,mutate", [
    # THE DEFECT ITSELF: the first N SHEET rows again.
    ("the tornado takes the first N sheet rows again",
     lambda src: src.replace(
         "                        f'=IF({eligible}=\"\",NA(),IF({source}=\"\",NA(),{source}))'))",
         "                        f'=IF({source}=\"\",NA(),{source})'))", 1)),
    # ELIGIBILITY INFERRED FROM THE BLANK RHO - the consequence read instead of
    # the cause, and a weaker second statement of the contract's rule.
    ("eligibility is inferred from a blank rho",
     lambda src: src.replace('TORNADO_ELIGIBILITY_FIELD = "rank"',
                             'TORNADO_ELIGIBILITY_FIELD = "rho"', 1)),
    # THE CATEGORY KEPT AND ONLY THE BAR SUPPRESSED - the residual this
    # settlement exists to remove.
    ("the category is kept and only the bar is gated",
     lambda src: src.replace(
         "            source = f\"{sheet}!${columns[str(column['source'])]}${source_row}\"",
         "            source = f\"{sheet}!${columns[str(column['source'])]}${source_row}\"\n"
         "            if str(column['source']) != 'rho':\n"
         "                out.append((row, str(column['column']),\n"
         "                            f'=IF({source}=\"\",NA(),{source})'))\n"
         "                continue", 1)),
    # THE EXCLUDED ROW COERCED TO SOMETHING RATHER THAN ABSENT.
    ("an excluded row is coerced to zero",
     lambda src: src.replace(
         "                        f'=IF({eligible}=\"\",NA(),IF({source}=\"\",NA(),{source}))'))",
         "                        f'=IF({eligible}=\"\",0,IF({source}=\"\",NA(),{source}))'))", 1)),
    # A MEASURED ZERO EXCLUDED - the opposite error, and just as wrong.
    ("a measured zero is excluded with the undefined ones",
     lambda src: src.replace(
         "                        f'=IF({eligible}=\"\",NA(),IF({source}=\"\",NA(),{source}))'))",
         "                        f'=IF({source}=0,NA(),IF({source}=\"\",NA(),{source}))'))", 1)),
    # THE ORDER RE-DERIVED IN RESULTS.
    ("the bridge re-ranks in Results",
     lambda src: src.replace(
         "            source = f\"{sheet}!${columns[str(column['source'])]}${source_row}\"",
         "            source = (f\"LARGE({sheet}!${columns[str(column['source'])]}\"\n"
         "                      f\"$13:${columns[str(column['source'])]}$22,{index + 1})\")",
         1)),
    # THE WINDOW GROWN PAST TEN.
    ("the top N exceeds ten",
     lambda src: src.replace("    for index in range(int(block[\"top_n\"])):",
                             "    for index in range(int(block[\"top_n\"]) + 5):", 1)),
])
def test_68_each_way_of_getting_the_tornado_input_wrong_is_refused(
        name: str, mutate) -> None:
    with pytest.raises((AssertionError, KeyError, ValueError)):
        _tornado_input_rules(_bridge_rows(mutate))


def test_69_the_tornado_input_rules_pass_on_the_real_builder() -> None:
    """SO THE SEVEN REFUSALS ABOVE ARE REFUSALS OF THE MUTATION, and the emitted
    rows really are what the built workbook holds."""
    import openpyxl
    from pccm_builder.spec_loader import load_spec
    rows = _bridge_rows()
    _tornado_input_rules(rows)
    charts = load_spec(SPEC / "workbook.yaml").phase6_shell["charts"]
    results = openpyxl.load_workbook(
        PCCM_ROOT / "build" / "PCCM_stageA.xlsx")[charts["bridge_sheet"]]
    for row, column, formula in rows:
        assert results[f"{column}{row}"].value == formula, (
            f"{column}{row}: the builder and the built workbook disagree")


@pytest.mark.parametrize("name,mutate", [
    # THE DIAGNOSTIC ROWS SIMPLY NOT WRITTEN. Excluding a driver from the chart
    # must never become excluding it from the record.
    ("the diagnostic records stop being appended",
     lambda src: src.replace(
         "        If results(index).Status <> SIM_SENSITIVITY_DEFINED Then\n"
         "            FillRecord block, slot + 1, results(index), 0\n"
         "            slot = slot + 1\n"
         "        End If", "", 1)),
    # THE STAMPED COUNT NARROWED TO THE RANKED ONES, which would blank the
    # diagnostic rows on the sheet through the count bound.
    ("the stamped count covers only the ranked drivers",
     lambda src: src.replace(
         "SIM_SENSITIVITY_STAMP_ROW_RECORD_COUNT).Value2 = driverCount",
         "SIM_SENSITIVITY_STAMP_ROW_RECORD_COUNT).Value2 = eligibleCount", 1)),
    # THE COMPLETENESS CHECK DROPPED, so a missing record would go unnoticed.
    ("the publication stops proving it wrote every driver",
     lambda src: src.replace("    If slot <> driverCount Then",
                             "    If False Then", 1)),
])
def test_69a_removing_the_diagnostic_rows_from_persistence_is_refused(
        name: str, mutate) -> None:
    """RETAINED DIAGNOSTICALLY IS A CONTRACT CLAUSE. The chart correction must
    not be paid for out of the record."""
    source = POST_BAS.read_text(encoding="utf-8")
    mutated = mutate(source)
    assert mutated != source, f"the mutation '{name}' changed nothing"
    publish = re.search(r"^Private Function Publish.*?^End Function",
                        mutated, re.S | re.M)
    assert publish, "the publication is gone"
    publish = publish.group(0)
    with pytest.raises(AssertionError):
        assert "If results(index).Status <> SIM_SENSITIVITY_DEFINED Then" in publish, (
            "the diagnostic records are no longer appended")
        assert ("SIM_SENSITIVITY_STAMP_ROW_RECORD_COUNT).Value2 = driverCount"
                in publish), "the stamped count no longer covers them"
        assert "If slot <> driverCount Then" in publish, (
            "the publication no longer proves it wrote a row for every driver")


def test_70_the_diagnostic_row_is_never_removed_from_persistence() -> None:
    """THE ONE THING THIS CORRECTION MAY NOT DO. Excluding a driver from the
    CHART is not removing it from the record - Publish still writes every
    driver, and the stamped count still covers them all."""
    publish = _procedure("Publish")
    assert "If results(index).Status <> SIM_SENSITIVITY_DEFINED Then" in publish
    assert "SIM_SENSITIVITY_STAMP_ROW_RECORD_COUNT).Value2 = driverCount" in publish
    assert "If slot <> driverCount Then" in publish, (
        "the publication no longer proves it wrote a row for every driver")
    # AND NOTHING IN THE BRIDGE BUILDER TOUCHES PERSISTENCE.
    from pccm_builder import workbook_builder as real
    body = re.search(r"^def _chart_bridge_drivers\(.*?(?=\n\ndef )",
                     Path(real.__file__).read_text(encoding="utf-8"), re.S | re.M)
    for reach in ("_SimData", "ClearRecords", "record_count", "Value2"):
        assert reach not in body.group(0), (
            f"the tornado bridge reaches into persistence: {reach}")
