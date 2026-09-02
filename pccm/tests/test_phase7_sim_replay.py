#!/usr/bin/env python3
"""PCCM Phase 7 Step-3 conformance tests for per-driver replay in `modSimEngine`.

Deterministic reconstruction of ONE driver's nominal contribution sequence for
an already-accepted run, and the single shared contribution routine both the
normal simulation and the replay reach.

--------------------------------------------------------------------------------
WHAT THESE TESTS PROVE, AND WHAT THEY DO NOT
--------------------------------------------------------------------------------
SOURCE CONFORMANCE, on Linux, now, through the accepted Phase-6 transcriber: the
replay reproduces the normal simulation's own numbers BIT FOR BIT, on models
where the comparison is exact and needs no tolerance at all.

VBA EXECUTION CONFORMANCE IS NOT PROVED. `modSimEngine` now carries bytes Run 6
never executed. Nothing here may be read as "VBA replayed a driver", and the
modified engine is NOT Windows/runtime accepted until a Phase-7 Windows run says
so.

WHY THE ORACLE NEEDS NO TOLERANCE. On a ONE-DRIVER model the accepted per-
iteration total IS that driver's contribution, so `SimEngineRun`'s own published
output is the reference and `==` is the honest comparison. On multi-driver models
the replayed columns are recombined with the accepted canonical accumulation and
compared to the same published total. Neither route consults a second
implementation of the arithmetic, which is the point of the extraction.

Runs standalone or under pytest.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder.calc_numeric import safe_signed_sum  # noqa: E402
from pccm_builder.vba_source import VbaModule  # noqa: E402

# THE ACCEPTED PHASE-6 ENGINE HARNESS, reused rather than rebuilt. A second
# transcription setup would be a second definition of what "the engine" is, and
# the whole subject of this package is not having two of those.
import test_phase6_sim_engine_vba as engine  # noqa: E402
from phase6_vba_transcribe import _Ref, _val  # noqa: E402

SIM_ENGINE_BAS = PCCM_ROOT / "src" / "vba" / "modSimEngine.bas"
SEED = 4242
N = 1000  # the accepted business minimum; the normal run refuses fewer


def _code() -> str:
    return VbaModule(name="modSimEngine", path=SIM_ENGINE_BAS,
                     raw=SIM_ENGINE_BAS.read_text(encoding="utf-8")).code


def _records(costs=(), risks=(), order=None):
    return engine._factor_records(engine._model(costs, risks), order=order)


def _totals(records, seed=SEED, iterations=N):
    ok, nominal, _pv, detail = engine._run(records, seed=seed, iterations=iterations)
    assert ok, detail
    return [_val(v) for v in nominal]


def _replay(records, permanent_id, seed=SEED, iterations=N):
    contributions, detail = [], _Ref("")
    ok = engine._transcribe()["SimEngineReplayDriver"](
        records, _Ref(len(records)), _Ref(seed), _Ref(iterations),
        _Ref(permanent_id), contributions, detail)
    assert ok, detail.v
    return [_val(v) for v in contributions]


def _replay_refuses(records, permanent_id, seed=SEED, iterations=N):
    contributions, detail = [], _Ref("")
    ok = engine._transcribe()["SimEngineReplayDriver"](
        records, _Ref(len(records)), _Ref(seed), _Ref(iterations),
        _Ref(permanent_id), contributions, detail)
    assert not ok, "the replay accepted an input it cannot honour"
    return detail.v


def _cost(pid, dist="Triangular", lo=80.0, ml=100.0, hi=130.0, quantity=2.0):
    return engine._cost(pid, dist, lo, ml, hi, quantity=quantity)


def _risk(pid, dist="Triangular", lo=100.0, ml=200.0, hi=400.0, probability=0.3):
    return engine._risk(pid, dist, lo, ml, hi, probability=probability)


# ===========================================================================
# A. EXACT REPLAY FIDELITY - one driver, so the total IS the contribution
# ===========================================================================
# Every case below is compared with `==`. There is no tolerance here and none is
# needed: both sides are produced by the same contribution routine from the same
# sampled input, so anything but bit equality would be a defect rather than a
# rounding difference.
SINGLE_DRIVER_CASES = (
    ("A degenerate cost line", [_cost("C-001", "Uniform", 50.0, None, 50.0)], [], "C-001"),
    ("B Uniform cost line", [_cost("C-001", "Uniform", 10.0, None, 20.0)], [], "C-001"),
    ("C Triangular cost line", [_cost("C-001", "Triangular", 80.0, 100.0, 130.0)], [], "C-001"),
    ("D Beta-PERT cost line", [_cost("C-001", "Beta-PERT", 80.0, 100.0, 130.0)], [], "C-001"),
    ("E Risk probability 0", [], [_risk("R-001", probability=0.0)], "R-001"),
    ("F Risk probability 1", [], [_risk("R-001", probability=1.0)], "R-001"),
    ("G Risk probability 0.2", [], [_risk("R-001", probability=0.2)], "R-001"),
    ("H Risk Uniform severity", [],
     [_risk("R-001", "Uniform", 100.0, None, 400.0, probability=0.4)], "R-001"),
    ("I Risk Triangular severity", [],
     [_risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=0.4)], "R-001"),
    ("J Risk Beta-PERT severity", [],
     [_risk("R-001", "Beta-PERT", 100.0, 200.0, 400.0, probability=0.4)], "R-001"),
)


def test_01_every_single_driver_case_replays_bit_for_bit() -> None:
    for label, costs, risks, target in SINGLE_DRIVER_CASES:
        records = _records(costs, risks)
        published = _totals(records)
        replayed = _replay(records, target)
        assert len(replayed) == N, (label, len(replayed))
        assert replayed == published, (
            f"{label}: the replayed contribution sequence is not the published one")


def test_01b_replay_is_bit_exact_where_nominal_and_pv_differ() -> None:
    """WITHOUT THIS THE MEASURE IS INVISIBLE. Every fixture above collapses to
    Knom = Kpv = 1, so a replay that emitted the PV contribution as nominal
    would agree with the published total anyway. Two applied years and a
    discount rate separate them, and the accepted engine battery keeps this
    fixture for the same reason.
    """
    model = engine._discounted_model(
        [engine._spread("CL-001", "cost", quantity=2.0)],
        [engine._spread("R-001", "risk", "Triangular", 100.0, 200.0, 400.0,
                        probability=0.4)])
    records = engine._factor_records(model)
    assert {r["Knom"] for r in records} == {1.0}
    assert all(0.0 < r["Kpv"] < 1.0 for r in records), "the fixture lost its discount"

    ok, nominal, pv, detail = engine._run(records, seed=SEED, iterations=N)
    assert ok, detail
    published = [_val(v) for v in nominal]
    assert published != [_val(v) for v in pv], "the two measures are indistinguishable"

    columns = [_replay(records, permanent_id) for permanent_id in ("CL-001", "R-001")]
    rebuilt = [safe_signed_sum([column[j] for column in columns], "engine")
               for j in range(N)]
    assert rebuilt == published, "the replay does not reproduce the NOMINAL total"
    # AND IT IS NOT THE PV TOTAL WEARING A NOMINAL LABEL.
    assert rebuilt != [_val(v) for v in pv]


def test_02_a_risk_at_probability_zero_contributes_zero_at_every_iteration() -> None:
    """And it is a sequence of observed zeros, not an absent driver."""
    records = _records([], [_risk("R-001", probability=0.0)])
    replayed = _replay(records, "R-001")
    assert len(replayed) == N
    assert set(replayed) == {0.0}


def test_03_a_risk_at_probability_one_contributes_at_every_iteration() -> None:
    records = _records([], [_risk("R-001", probability=1.0)])
    replayed = _replay(records, "R-001")
    assert 0.0 not in replayed, "a certain risk failed to occur"


# ===========================================================================
# B. THE SEVERITY STREAM ADVANCES ON EVERY ITERATION
# ===========================================================================
def test_04_severity_advances_on_non_occurrence_iterations() -> None:
    """D6-18b, proved by the only evidence that can distinguish the two designs.

    At probability 1 every iteration occurs, so contribution / Knom IS the
    severity sequence. At probability 0.2 the same stream must produce the SAME
    severity at the SAME iteration index - which can only be true if the
    sampler was invoked on the ~80% of iterations where the risk did not occur.

    The falsifier is asserted too: under a CONDITIONAL stream the occurring
    iterations would carry the first k severities in order, and they do not.
    """
    def records_at(probability):
        return _records([], [_risk("R-001", "Beta-PERT", 100.0, 200.0, 400.0,
                                   probability=probability)])

    certain, partial = records_at(1.0), records_at(0.2)
    knom = certain[0]["Knom"]
    assert knom == partial[0]["Knom"], "the factor moved with the probability"

    severity = [value / knom for value in _replay(certain, "R-001")]
    contributions = _replay(partial, "R-001")
    occurred = [j for j, value in enumerate(contributions) if value != 0.0]
    assert 0 < len(occurred) < N, len(occurred)

    for j in occurred:
        assert contributions[j] / knom == severity[j], (
            f"iteration {j} carries a severity the unconditional stream would not "
            "have produced there")

    # UNDER A CONDITIONAL STREAM the k-th occurrence would carry severity[k].
    conditional = [severity[k] for k in range(len(occurred))]
    assert not all(contributions[j] / knom == conditional[k]
                   for k, j in enumerate(occurred)), (
        "the occurring iterations match a CONDITIONAL severity path, so this "
        "control cannot tell the two designs apart")


def test_05_probability_changes_which_iterations_occur_and_not_the_severities() -> None:
    """The property D6-18b exists for: two runs differing only in one
    Probability stay comparable, because the severity path did not move."""
    def severities(probability):
        records = _records([], [_risk("R-001", "Triangular", 100.0, 200.0, 400.0,
                                      probability=probability)])
        knom = records[0]["Knom"]
        return records, knom, _replay(records, "R-001")

    _r_hi, knom_hi, high = severities(0.9)
    _r_lo, knom_lo, low = severities(0.1)
    assert knom_hi == knom_lo
    shared = [j for j in range(N) if high[j] != 0.0 and low[j] != 0.0]
    assert shared, "no iteration occurred under both probabilities"
    for j in shared:
        assert high[j] == low[j], (
            f"iteration {j} drew a different severity when only Probability changed")


# ===========================================================================
# C. PER-ITERATION TOTAL RECONCILIATION
# ===========================================================================
def test_06_replaying_every_driver_rebuilds_the_published_total() -> None:
    """THE JOINT PROOF. Each driver replayed separately, recombined with the
    ACCEPTED canonical accumulation - `SafeSignedSum` over the whole vector in
    canonical order, never a running total - equals the published TotalNom.

    This is verification, not a second production total: nothing in `src/vba`
    accumulates this way, and the recombination lives here.
    """
    records = _records(
        [_cost("C-002", "Uniform", 10.0, None, 20.0, quantity=2.0),
         _cost("C-005", "Beta-PERT", 80.0, 100.0, 130.0, quantity=1.5)],
        [_risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=0.2),
         _risk("R-003", "Uniform", 50.0, None, 90.0, probability=0.75)])
    published = _totals(records)
    # CANONICAL ORDER: every cost line in ordinal permanent-id order, then every
    # risk in ordinal permanent-id order.
    order = ["C-002", "C-005", "R-001", "R-003"]
    columns = [_replay(records, permanent_id) for permanent_id in order]
    rebuilt = [safe_signed_sum([column[j] for column in columns], "engine")
               for j in range(N)]
    assert rebuilt == published, (
        f"{sum(1 for j in range(N) if rebuilt[j] != published[j])} iterations "
        "disagree with the published total")


def test_07_the_reconciliation_would_notice_a_missing_driver() -> None:
    """A control that passes with a driver left out is not a reconciliation."""
    records = _records([_cost("C-002", "Uniform", 10.0, None, 20.0)],
                       [_risk("R-001", probability=0.5)])
    published = _totals(records)
    partial = [_replay(records, "C-002")]
    rebuilt = [safe_signed_sum([column[j] for column in partial], "engine")
               for j in range(N)]
    assert rebuilt != published


# ===========================================================================
# D. ITERATION IDENTITY
# ===========================================================================
def test_08_the_vector_is_one_observation_per_iteration_in_order() -> None:
    """No sorting, no compaction of non-occurrence zeros, no renumbering. P7-4
    pairs this positionally with the persisted TotalNom."""
    records = _records([], [_risk("R-001", probability=0.2)])
    replayed = _replay(records, "R-001")
    assert len(replayed) == N, "observations were dropped or added"
    assert replayed.count(0.0) > 0, "the non-occurrence zeros were compacted away"
    assert replayed != sorted(replayed), "the vector came back sorted"
    # AND IT IS REPRODUCIBLE: the same request gives the same vector.
    assert _replay(records, "R-001") == replayed


def test_09_a_shorter_replay_is_a_prefix_of_a_longer_one() -> None:
    """Iteration j means the same thing whatever N is asked for, which is what
    makes positional pairing meaningful."""
    records = _records([_cost("C-001", "Beta-PERT")], [])
    long_run = _replay(records, "C-001", iterations=N)
    short_run = _replay(records, "C-001", iterations=N // 2)
    assert short_run == long_run[: N // 2]


# ===========================================================================
# E. UNRELATED-DRIVER ISOLATION - exactly what the stream contract guarantees
# ===========================================================================
# MEASURED, NOT ASSUMED. Component stream indices come from a canonical sequence
# ordered by kind, permanent id and role, so a driver inserted EARLIER in that
# sequence shifts every stream after it. That is a property of the accepted
# Phase-6 contract, not a defect, and claiming isolation against it would be
# claiming something the contract does not provide. What the contract does
# guarantee is asserted here, and the boundary is asserted with it.
def test_10_supply_order_is_not_driver_identity() -> None:
    """The strongest of these, and the one sensitivity depends on: handing the
    engine the same drivers in a different PHYSICAL order changes nothing,
    because ordering is canonical and a worksheet row is not identity."""
    costs = [_cost("C-002"), _cost("C-004")]
    risks = [_risk("R-001")]
    baseline = _records(costs, risks)
    reversed_supply = _records(costs, risks, order=[1, 0, 2])
    for target in ("C-002", "C-004", "R-001"):
        assert _replay(reversed_supply, target) == _replay(baseline, target), target


def test_11_a_canonically_later_driver_does_not_disturb_an_earlier_one() -> None:
    baseline = _records([_cost("C-002"), _cost("C-004")], [_risk("R-001")])
    before = _replay(baseline, "C-002")
    appended = _records([_cost("C-002"), _cost("C-004"), _cost("C-009")],
                        [_risk("R-001")])
    assert _replay(appended, "C-002") == before, "an appended cost line moved C-002"
    removed = _records([_cost("C-002")], [_risk("R-001")])
    assert _replay(removed, "C-002") == before, "deleting a later cost line moved C-002"
    # A LATER RISK disturbs neither the cost line nor the earlier risk.
    risk_added = _records([_cost("C-002"), _cost("C-004")],
                          [_risk("R-001"), _risk("R-007")])
    assert _replay(risk_added, "C-002") == before
    assert _replay(risk_added, "R-001") == _replay(baseline, "R-001")


def test_12_the_isolation_boundary_is_stated_rather_than_overclaimed() -> None:
    """A driver inserted EARLIER in the canonical sequence DOES move the streams
    after it, and every risk sits after every cost line. Asserting the change is
    what keeps the guarantee above honest: it says which additions are safe by
    showing which are not.
    """
    baseline = _records([_cost("C-002"), _cost("C-004")], [_risk("R-001")])
    earlier = _records([_cost("C-001"), _cost("C-002"), _cost("C-004")],
                       [_risk("R-001")])
    assert _replay(earlier, "C-002") != _replay(baseline, "C-002"), (
        "an earlier cost line left C-002's stream index unchanged; the canonical "
        "assignment is not what this control believes it to be")
    assert _replay(earlier, "R-001") != _replay(baseline, "R-001"), (
        "adding a cost line left the risk streams unchanged; risks are assigned "
        "after every cost component and must have moved")


# ===========================================================================
# F. REPLAY IS OBSERVATIONAL, AND OWNS NOTHING
# ===========================================================================
def test_13_replay_refuses_what_it_cannot_honour() -> None:
    records = _records([_cost("C-001")], [])
    assert "not in this model" in _replay_refuses(records, "C-999")
    assert "at least one iteration" in _replay_refuses(records, "C-001", iterations=0)


def test_14_the_replay_reaches_no_state_identity_or_workbook() -> None:
    body = engine._procedure("SimEngineReplayDriver")
    stripped = VbaModule(name="probe", path=SIM_ENGINE_BAS, raw=body).code
    for token in ("run_id", "RunId", "Nonce", "F21", "_SimData", "Worksheet",
                  "Range(", "Cells(", "ThisWorkbook", "Application.", "MsgBox",
                  "SimReport", "Digest", "Attempt"):
        assert token not in stripped, f"the replay reaches {token!r}"


def test_15_the_replay_advances_only_the_target_drivers_streams() -> None:
    """Preparation derives every component's initial state by jump-ahead and
    draws nothing; only the target's states are then advanced."""
    body = VbaModule(name="probe", path=SIM_ENGINE_BAS,
                     raw=engine._procedure("SimEngineReplayDriver")).code
    for sampler in ("SimSampleBernoulli", "SimEngineSampleValue"):
        calls = re.findall(rf"{sampler}\(([^,]+),", body)
        assert calls, sampler
        for first_argument in calls:
            assert "prepared(target)" in first_argument or "occurrenceState" in first_argument \
                or "valueState" in first_argument, (sampler, first_argument)
    # NO INDEXED STATE ARRAYS: the normal loop holds one state per driver, the
    # replay holds exactly two scalars, so there is nothing else it could advance.
    assert "valueState(" not in body and "occurrenceState(" not in body


def test_16_the_replay_is_sequential_and_seeks_nothing() -> None:
    body = VbaModule(name="probe", path=SIM_ENGINE_BAS,
                     raw=engine._procedure("SimEngineReplayDriver")).code
    assert re.search(r"For iteration = 1 To iterations", body), (
        "the replay does not advance iteration by iteration")
    for seek in ("Jump", "Skip", "Seek", "Advance("):
        assert seek not in body, f"the replay attempts a {seek!r}"


# ===========================================================================
# G. ONE CONTRIBUTION ROUTINE, AND BOTH PATHS REACH IT
# ===========================================================================
def test_17_the_contribution_arithmetic_has_exactly_one_owner() -> None:
    """THE POINT OF THE EXTRACTION. If the replay carried its own copy of this
    expression the explanation and the published number could drift apart while
    both looked right."""
    code = _code()
    assert code.count("Private Function SimEngineContribution(") == 1
    # The products themselves appear ONLY inside that routine.
    body = VbaModule(name="probe", path=SIM_ENGINE_BAS,
                     raw=engine._procedure("SimEngineContribution")).code
    outside = code.replace(body, "")
    assert "SafeProduct(" not in outside, (
        "a second place forms a contribution product")
    # AND QUANTITY IS APPLIED IN EXACTLY ONE PLACE. It reaches the module twice
    # more - copied in during adoption and range-checked during validation -
    # and neither multiplies by it. Naming the two legitimate readers is what
    # makes this an assertion rather than a word count.
    readers = {line.strip() for line in outside.splitlines() if ".Quantity" in line}
    assert readers == {"target.Quantity = factor.Quantity",
                       "If Not IsUsableDouble(factor.Quantity) Then"}, sorted(readers)


def test_18_both_the_simulation_and_the_replay_call_it() -> None:
    for procedure, expected in (("SimEngineRun", 4), ("SimEngineReplayDriver", 1)):
        body = VbaModule(name="probe", path=SIM_ENGINE_BAS,
                         raw=engine._procedure(procedure)).code
        found = body.count("SimEngineContribution(")
        assert found == expected, (procedure, found, expected)


def test_19_the_shared_routine_keeps_the_accepted_shapes() -> None:
    body = VbaModule(name="probe", path=SIM_ENGINE_BAS,
                     raw=engine._procedure("SimEngineContribution")).code
    # A cost line multiplies three factors, a risk two, and a risk that did not
    # occur contributes exactly zero.
    assert "factors(1) = prepared.Quantity" in body
    assert "count = 3" in body and "count = 2" in body
    assert re.search(r"If Not occurred Then", body)
    # PROBABILITY IS SPENT ON THE BERNOULLI DRAW and appears in no factor here.
    assert ".Probability" not in body


def test_20_pv_is_not_derived_from_the_nominal_term() -> None:
    """The caller passes the other factor and gets an independent product."""
    body = VbaModule(name="probe", path=SIM_ENGINE_BAS,
                     raw=engine._procedure("SimEngineRun")).code
    assert "prepared(index).Kpv" in body and "prepared(index).Knom" in body
    for forbidden in ("* discount", "/ discount", "nominalTerm(index) *"):
        assert forbidden not in body, forbidden


# ===========================================================================
# H. THE P7-2 KERNEL STAYS PURE
# ===========================================================================
def test_21_the_sensitivity_kernel_acquired_no_replay_or_rng() -> None:
    """P7-3 is the bridge; the kernel is still only the mathematics."""
    kernel = PCCM_ROOT / "src" / "vba" / "modSimSensitivity.bas"
    code = VbaModule(name="modSimSensitivity", path=kernel,
                     raw=kernel.read_text(encoding="utf-8")).code
    for token in ("SimRng", "SimSample", "SimEngine", "Replay", "InitialState",
                  "Bernoulli", "Stream", "DriverFactors"):
        assert token not in code, f"the pure kernel acquired {token!r}"
    assert "SimEngineReplayDriver" not in code
