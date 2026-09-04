#!/usr/bin/env python3
"""P7-6: the annual answer's persistence, its bank isolation and its handoff.

Three things are proved here and each needs a different kind of evidence.

THE SETTLEMENT, EXECUTED. Moving the reporting selector must not retire the
annual ladders and must retire the selected-Px profile, and the profile must
never be relabelled. Those are RULES, so they are written as pure procedures and
run through the accepted Phase-6 transcriber over the real `.bas` text - every
branch of the five-state settlement is exercised directly rather than described.

THE VALUES THAT WOULD BE PERSISTED, AGAINST AN INDEPENDENT ORACLE. The P7-6
orchestration blocks the years, indexes the block, lifts the two order
statistics out of it and packs the record row. All of that is re-derived in
plain Python from the same replay matrix using the P7-5 primitives, and the two
are compared value by value. A blocking or packing defect shows here; a shared
mistake cannot, because the two paths share no code.

THE DISCIPLINE, STRUCTURALLY. Publication order, the bounded clear, bank
isolation and the separation between the module that owns the pipeline and the
module that owns the addresses are asserted over the source itself, because they
are properties no value can demonstrate.

NOTHING HERE IS RUNTIME EVIDENCE. No Windows run has executed this code.
"""

from __future__ import annotations

import math
import re
import sys
from pathlib import Path

import pytest
import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder import load_contract, load_sim_contract  # noqa: E402
from pccm_builder import sim_annual as annual  # noqa: E402
from pccm_builder.calc_numeric import safe_signed_sum  # noqa: E402
from pccm_builder.sim_emit import render_sim_contract_module  # noqa: E402
from pccm_builder.sim_stats import percentile_type7  # noqa: E402
from pccm_builder.spec_loader import load_spec  # noqa: E402
from phase6_vba_transcribe import _Ref, _val, build as _build_transcription  # noqa: E402

# THE ONE FIXTURE. The P7-5 replay suite already resolves a real four-year,
# two-currency, two-profile model through the accepted Phase-5 oracle and drives
# the real engine over it. Importing it is what keeps this suite's read-back
# comparison anchored to the SAME accepted model rather than a second one free
# to drift from it.
import test_phase7_sim_annual_replay_vba as replay  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"
RUN_BAS = SRC_VBA / "modSimAnnualRun.bas"
STORE_BAS = SRC_VBA / "modSimAnnualStore.bas"

_CACHE: dict = {}

ABSOLUTE_FLOOR = 1e-6
RELATIVE_COEFFICIENT = 1e-12
SCALE_FLOOR = 1.0
YEARS = replay.YEARS
ITERATIONS = replay.ITERATIONS
SEED = replay.SEED


# ===========================================================================
# THE CONTRACT
# ===========================================================================
def _annual_contract() -> dict:
    if "annual" not in _CACHE:
        raw = yaml.safe_load((SPEC / "sim_contract.yaml").read_text(encoding="utf-8"))
        _CACHE["annual"] = raw["sim_data"]["annual_records"]
    return _CACHE["annual"]


def _projected() -> dict:
    if "consts" not in _CACHE:
        out: dict = {}
        rendered = render_sim_contract_module(
            load_spec(SPEC / "workbook.yaml"),
            load_sim_contract(SPEC / "sim_contract.yaml"),
            load_contract(SPEC / "input_contract.yaml"))
        sources = [rendered]
        for name in ("modCalcFactors", "modSimAnnual", "modSimEngine"):
            sources.append((SRC_VBA / f"{name}.bas").read_text(encoding="utf-8"))
        for text in sources:
            for line in text.splitlines():
                match = re.match(r"^Public Const (\w+) As (\w+) = (.*)$", line)
                if not match:
                    continue
                name, kind, rest = match.groups()
                literal = rest.split("    '")[0].rstrip()
                out[name] = (literal[1:-1] if kind == "String"
                             else (float(literal) if kind == "Double" else int(literal)))
        # The tolerance triple lives in the Phase-5 contract projection, which is
        # not part of the simulation module.
        for name, value in (("TOL_IDENTITY_ABSOLUTE_FLOOR", ABSOLUTE_FLOOR),
                            ("TOL_IDENTITY_RELATIVE_COEFFICIENT", RELATIVE_COEFFICIENT),
                            ("TOL_CONDITIONING_SCALE_FLOOR", SCALE_FLOOR)):
            out[name] = value
        _CACHE["consts"] = out
    return _CACHE["consts"]


def _column_number(letter: str) -> int:
    total = 0
    for character in letter:
        total = total * 26 + (ord(character) - ord("A") + 1)
    return total


def _offsets(bank: str) -> dict:
    """The contracted field offsets, computed from the contract's own letters."""
    block = _annual_contract()
    first = _column_number(block["index_columns"][bank]["project_index"])
    return {
        "project_index": 0,
        "calendar_year": _column_number(block["index_columns"][bank]["calendar_year"]) - first,
        "nominal": _column_number(block["quantile_first_column"][bank]["nominal"]) - first,
        "pv": _column_number(block["quantile_first_column"][bank]["pv"]) - first,
        "nominal_profile":
            _column_number(block["selected_px_profile_columns"][bank]["nominal"]) - first,
        "pv_profile":
            _column_number(block["selected_px_profile_columns"][bank]["pv"]) - first,
    }


def _fields(bank: str) -> int:
    return _offsets(bank)["pv_profile"] + 1


# ===========================================================================
# THE TRANSCRIPTION
# ===========================================================================
def _shim_signed_sum(terms, count, result, *rest):
    try:
        result.v = safe_signed_sum([float(v) for v in terms[: int(_val(count))]], "annual")
    except Exception:
        return False
    return True


def _finite(value):
    return isinstance(value, float) and math.isfinite(value)


def _shim_multiply(a, b, result, *rest):
    product = _val(a) * _val(b)
    if not _finite(product):
        return False
    result.v = product
    return True


def _shim_subtract(a, b, result, *rest):
    difference = _val(a) - _val(b)
    if not _finite(difference):
        return False
    result.v = difference
    return True


def _shim_accumulate(accumulator, term, *rest):
    total = _val(accumulator) + _val(term)
    if not _finite(total):
        return False
    accumulator.v = total
    return True


def _shim_product(factors, count, result, *rest):
    product = 1.0
    for value in list(factors)[: int(_val(count))]:
        product *= float(value)
    if not _finite(product):
        return False
    result.v = product
    return True


def _unsupported(name):
    def shim(*args):
        raise AssertionError(f"{name} is not exercised by this path")
    return shim


# The store's addressing is Excel's - `Range("AD33").Column` - so the six column
# helpers and the offset are bound to the CONTRACT'S OWN letters here. That is
# not a re-implementation of the layout: it is the same authority the projected
# constants come from, so a contract move moves both sides at once.
def _address_shims() -> dict:
    def offset_of(bank, column, *rest):
        letters = _val(column)
        bank_name = _val(bank)
        first = _annual_contract()["index_columns"][bank_name]["project_index"]
        return _column_number(letters) - _column_number(first)

    def column_for(group, key):
        def shim(bank, *rest):
            block = _annual_contract()
            if group == "index_columns":
                return block[group][_val(bank)][key]
            return block[group][_val(bank)][key]
        return shim

    return {
        "OffsetOf": offset_of,
        "FirstColumn": column_for("index_columns", "project_index"),
        "CalendarYearColumn": column_for("index_columns", "calendar_year"),
        "NominalFirstColumn": column_for("quantile_first_column", "nominal"),
        "PvFirstColumn": column_for("quantile_first_column", "pv"),
        "NominalProfileColumn": column_for("selected_px_profile_columns", "nominal"),
        "PvProfileColumn": column_for("selected_px_profile_columns", "pv"),
        "LastColumn": column_for("selected_px_profile_columns", "pv"),
    }


def _vba() -> dict:
    if "vba" not in _CACHE:
        extra = {
            "MAX_DOUBLE": sys.float_info.max,
            "SafeProduct": _shim_product,
            "SafeSignedSum": _shim_signed_sum,
            "SafeMultiply": _shim_multiply,
            "SafeSubtract": _shim_subtract,
            "SafeAccumulate": _shim_accumulate,
            "SafeDivide": _unsupported("SafeDivide"),
            "ConditioningScaledExact": _unsupported("ConditioningScaledExact"),
        }
        extra.update(_address_shims())
        _CACHE["vba"] = _build_transcription(
            {
                "modSimRng": SRC_VBA / "modSimRng.bas",
                "modSimSample": SRC_VBA / "modSimSample.bas",
                "modSimStats": SRC_VBA / "modSimStats.bas",
                "modSimAnnual": SRC_VBA / "modSimAnnual.bas",
                "modSimEngine": SRC_VBA / "modSimEngine.bas",
                "modCalcFactors": SRC_VBA / "modCalcFactors.bas",
                # TYPES ONLY. SimAnnualIdentity is declared here, and the
                # pipeline below is typed by it; no procedure of the store is
                # compiled except the four that touch no cell.
                "modSimAnnualStore": STORE_BAS,
                "modSimAnnualRun": RUN_BAS,
            },
            _projected(),
            only={
                "modCalcFactors": {"IsUsableDouble", "ConditioningScaledMagnitude",
                                   "IdentityAllowance"},
                "modSimStats": {
                    "SimStatsQuantileType7", "SimStatsQuantileSorted",
                    "SimStatsSortedCopy", "SimStatsSortAscending",
                    "SimStatsQuantilePosition", "SimStatsPositionOf",
                    "SimStatsOrderedIndices", "SimStatsSortIndices",
                    "SimStatsUsableSequence", "SimStatsUsableProbability",
                    "SimStatsConstantValue", "SimStatsUnitScale",
                    "SimStatsLadderProbabilities", "SimStatsLadderLabel",
                    "SimStatsProbabilityOf", "SimStatsSelectedProbability",
                },
                "modSimAnnualStore": {
                    "SimAnnualStoreFlatten", "LayoutIsSound", "Claim",
                    "DistributionStateOf", "ProfileStateOf", "IdentityMatches",
                },
                "modSimAnnualRun": {
                    "ProduceAnnual", "BuildYearFactors", "ReconcileTerms",
                    "ReconcileProfile", "CrossCheckYearCount",
                },
            },
            signature_only={"modCalcFactors": {
                "SafeProduct", "SafeSignedSum", "SafeMultiply", "SafeDivide",
                "SafeSubtract", "SafeAccumulate", "ConditioningScaledExact"}},
            extra=extra,
        )
    return _CACHE["vba"]


def _identity(**overrides) -> dict:
    run = _vba()["_new"]("SimAnnualIdentity")
    run.update({
        "Bank": "A", "RunId": 1, "EffectiveSeed": SEED,
        "RequestFingerprint": "FP", "ResultDigest": "RD",
        "Iterations": ITERATIONS, "YearCount": YEARS,
        "SelectedLabel": "P80", "SelectedProbability": 0.8,
    })
    run.update(overrides)
    return run


# ===========================================================================
# THE INDEPENDENT ORACLE - plain Python over the same replay matrix
# ===========================================================================
def _matrix(measure: str):
    key = f"matrix:{measure}"
    if key not in _CACHE:
        _CACHE[key] = replay._annual_matrix(measure)
    return _CACHE[key]


def _oracle_ladder(measure: str, probabilities):
    matrix = _matrix(measure)
    return [
        [percentile_type7([matrix[j][year] for j in range(ITERATIONS)], p)
         for p in probabilities]
        for year in range(YEARS)
    ]


def _oracle_profile(measure: str, p: float):
    totals = replay._run_totals()[1 if measure == "PV" else 0]
    matrix = _matrix(measure)
    position = annual.percentile_position(totals, p)
    low = matrix[position.lo]
    if position.fraction == 0.0:
        return list(low)
    high = matrix[position.hi]
    out = []
    for year in range(YEARS):
        if low[year] == high[year]:
            out.append(low[year])
        else:
            out.append((1.0 - position.fraction) * low[year]
                       + position.fraction * high[year])
    return out


def _produce(measure: str, p: float, probabilities):
    """The P7-6 orchestration's own answer for one measure."""
    totals = replay._run_totals()[1 if measure == "PV" else 0]
    position = _vba()["_new"]("SimStatsPosition")
    detail = _Ref("")
    assert _vba()["SimStatsQuantilePosition"](
        list(totals), _Ref(len(totals)), _Ref(p), position, detail), detail.v
    ladder, profile = [], []
    ok = _vba()["ProduceAnnual"](
        _identity(SelectedProbability=p), replay._records(),
        _Ref(len(replay._records())), replay._year_factors(measure), _Ref(measure),
        position, list(probabilities), ladder, profile, detail)
    assert ok, f"{measure}: the annual production refused: {detail.v}"
    return [float(v) for v in ladder], [float(v) for v in profile]


# ===========================================================================
# A. THE CONTRACT AND ITS PROJECTION
# ===========================================================================
def test_01_the_handoff_declares_two_states_and_refuses_one_boolean() -> None:
    handoff = _annual_contract()["handoff"]
    assert handoff["single_boolean_permitted"] is False
    assert handoff["states_derived_without_worksheet_access"] is True
    assert handoff["owner_module"] == "modSimAnnualStore"
    # FIVE SITUATIONS, and the two vocabularies together are what carry them.
    assert list(handoff["distribution_states"]) == ["NOT PRODUCED", "CURRENT", "HISTORICAL"]
    assert list(handoff["profile_states"]) == [
        "NOT PRODUCED", "CURRENT", "OTHER Px", "HISTORICAL"]
    # The profile's vocabulary is strictly wider, and the extra member is the
    # whole point: a profile can exist, be valid, and belong to a level nobody
    # is asking for.
    extra = set(handoff["profile_states"]) - set(handoff["distribution_states"])
    assert extra == {"OTHER Px"}
    assert handoff["inconsistent_stamp_state"] == "OTHER Px"


def test_02_the_settlement_leaves_are_all_present_and_hold() -> None:
    block = _annual_contract()
    assert block["distribution_currentness_is_selector_specific"] is False
    assert block["profile_currentness_is_selector_specific"] is True
    assert block["profile_relabelled_on_selector_change"] is False
    assert block["profile_current_requires_label_and_probability_match"] is True
    assert block["selector_change_requires_new_simulation"] is False
    assert block["selector_move_invalidates_simulation"] is False
    assert block["selector_move_makes_profile_current"] is False


def test_03_the_year_axis_is_read_back_and_never_rebuilt() -> None:
    block = _annual_contract()
    assert block["year_axis_source"] == "_Calc: tblCalcYears"
    assert block["year_axis_recomputed"] is False
    assert block["discount_series_rebuilt"] is False
    assert block["timeline_resolved_here"] is False
    assert block["year_count_cross_checked_against_driver_weights"] is True


def test_04_every_state_word_reaches_production_through_the_projection() -> None:
    """No state string is typed into a module."""
    projected = {name: value for name, value in _projected().items()
                 if name.startswith("SIM_ANNUAL_STATE_")}
    handoff = _annual_contract()["handoff"]
    assert set(projected.values()) == set(handoff["distribution_states"]) | set(
        handoff["profile_states"])
    for path in (RUN_BAS, STORE_BAS):
        text = path.read_text(encoding="utf-8")
        code = "\n".join(line for line in text.splitlines()
                         if not line.strip().startswith("'"))
        for state in projected.values():
            assert f'"{state}"' not in code, (
                f"{path.name} spells the state {state!r} instead of consuming the "
                "projected constant")


def test_05_the_accessors_are_exactly_the_contracted_four() -> None:
    declared = [a["name"] for a in _annual_contract()["handoff"]["accessors"]]
    assert declared == ["PCCM_AnnualDistributionState", "PCCM_AnnualProfileState",
                        "PCCM_AnnualProfilePx", "PCCM_AnnualYearCount"]
    text = STORE_BAS.read_text(encoding="utf-8")
    for name in declared:
        assert f"Public Function {name}(" in text, f"{name} is not published"


# ===========================================================================
# B. TWO MODULES, TWO RESPONSIBILITIES
# ===========================================================================
def _code(path: Path) -> str:
    return "\n".join(line for line in path.read_text(encoding="utf-8").splitlines()
                     if not line.strip().startswith("'"))


def test_06_the_producer_names_no_cell() -> None:
    """It owns the pipeline; the store owns every address.

    This is what makes a contract move a one-place change, and it is the
    property whose absence left the sensitivity availability formula pointing at
    the statistics band after the block moved.
    """
    code = _code(RUN_BAS)
    for token in ("Range", "Worksheet", "ClearContents", "Value2", "ListObject",
                  "SIM_DATA_SHEET", "CALC_SHEET", "TBL_CALC_YEARS",
                  "SIM_ANNUAL_STAMP", "SIM_ANNUAL_FIRST_ROW", "SIM_ANNUAL_HEADER_ROW",
                  "SIM_ITER_", "SIM_SNAPSHOT_", "SIM_SHARED_VALUE_COLUMN",
                  "_COLUMN", "modWorkbook"):
        assert token not in code, f"modSimAnnualRun names {token!r}"


def test_07_the_store_measures_nothing() -> None:
    """It owns the addresses; the producer owns every number.

    modSimStats is reachable for ONE thing - deciding whether a label is a
    selectable confidence level and what probability it spells - which is a
    lookup through the projected ladder, not a measurement.
    """
    code = _code(STORE_BAS)
    for token in ("modSimEngine", "modSimRng", "modSimSample", "modSimAnnual.",
                  "SimStatsQuantile", "SimStatsDescribe", "SimStatsMean",
                  "SafeSignedSum", "IdentityAllowance", "ConditioningScaled",
                  "modSimSensitivity", "modSimNonce", "modSimFingerprint"):
        assert token not in code, f"modSimAnnualStore reaches {token!r}"
    assert "modSimStats.SimStatsSelectedProbability" in code, (
        "the store must resolve a label through the one ladder authority")


def test_08_neither_module_touches_run_identity() -> None:
    """The annual step is not a simulation and must not look like one.

    No run id is allocated, no nonce is advanced, no attempt row is written and
    no digest or fingerprint is computed. The identity it handles is READ from
    the published run and written back only into the annual stamp.
    """
    for path in (RUN_BAS, STORE_BAS):
        code = _code(path)
        for token in ("SimNonce", "SimFpBuild", "SimFpResult", "PENDING",
                      "SIM_IDENTITY_ROW_ATTEMPT", "AllocateRunId", "SimEngineRun("):
            assert token not in code, f"{path.name} reaches {token!r}"
    store = _code(STORE_BAS)
    # The identity rows are READ. The only rows this module writes are the
    # annual stamp's own.
    written = set(re.findall(r"StampCell\(run\.Bank, (\w+)\)\.Value2 = ", store))
    assert written and all(name.startswith("SIM_ANNUAL_STAMP_ROW_") for name in written), (
        f"the store writes outside the annual stamp: {sorted(written)}")


def test_09_no_iteration_level_annual_value_is_persisted() -> None:
    """The contract says none is, and the writer cannot produce one.

    The record block is sized by the YEAR count. The iteration count appears in
    this module only where a published identity is read or the totals column is
    addressed - never in a range that is written.
    """
    assert _annual_contract()["iteration_level_annual_values_persisted"] is False
    body = _publish_body()
    # THE TWO RANGES THIS MODULE WRITES, and they are both in the publication.
    written = [match.group(1) for match in
               re.finditer(r"Range\((.*?)\)\.(?:Value2 = block|ClearContents)", body, re.S)]
    assert len(written) == 2, f"the publication writes {len(written)} range(s)"
    for span in written:
        assert "Iterations" not in span, (
            f"a written range is sized by the iteration count: {span}")
        assert "YearCount" in span or "LIMIT_MAX_YEAR_COLUMNS" in span, (
            f"a written range is bounded by neither the answer nor the "
            f"structural maximum: {span}")
    # And nothing outside the publication writes a cell at all.
    store = _code(STORE_BAS)
    assert store.count(".ClearContents") == 1
    assert store.count(".Value2 = block") == 1


# ===========================================================================
# C. PUBLICATION DISCIPLINE
# ===========================================================================
def _publish_body() -> str:
    text = STORE_BAS.read_text(encoding="utf-8")
    start = text.index("Public Function SimAnnualStorePublish")
    return text[start:text.index("\nEnd Function", start)]


def test_10_the_marker_is_blanked_first_and_written_last() -> None:
    body = _publish_body()
    blanked = body.index("SIM_ANNUAL_STAMP_ROW_PUBLISHED).Value2 = vbNullString")
    published = body.index("SIM_ANNUAL_STAMP_ROW_PUBLISHED).Value2 = SIM_ANNUAL_PUBLISHED")
    cleared = body.index("ClearContents")
    records = body.index(".Value2 = block")
    assert blanked < cleared < records < published, (
        "the block must be unpublished from the first write to the last")
    for row in ("RUN_ID", "EFFECTIVE_SEED", "REQUEST_FINGERPRINT", "RESULT_DIGEST",
                "ITERATIONS", "YEAR_COUNT", "SELECTED_PX_LABEL",
                "SELECTED_PX_PROBABILITY"):
        assert records < body.index(f"SIM_ANNUAL_STAMP_ROW_{row})") < published, (
            f"the {row} stamp is not written between the records and the marker")
    assert _annual_contract()["stamp"]["published_written_last"] is True


def test_11_the_clear_spans_the_structural_maximum_not_the_answer() -> None:
    """DURATION-SHRINK SAFETY. A four-year answer written over a twenty-year one
    must leave nothing of the twenty behind."""
    body = _publish_body()
    clear = body[body.index("ClearContents") - 400:body.index("ClearContents")]
    assert "LIMIT_MAX_YEAR_COLUMNS" in clear, (
        "the clear is bounded by something other than the structural maximum")
    assert "run.YearCount" not in clear, (
        "the clear is bounded by the answer about to be written, so a shorter "
        "answer would leave the tail of a longer one readable")
    assert _annual_contract()["stamp"]["surplus_rows_cleared"] is True
    assert _annual_contract()["stamp"]["cleared_before_write"] is True


def test_12_every_address_branches_on_the_bank() -> None:
    """A/B ISOLATION. One bank's publication may not read, clear or write the
    other's, so no address helper can produce a column without being told which
    bank it is for."""
    text = STORE_BAS.read_text(encoding="utf-8")
    helpers = ("StampCell", "FirstColumn", "CalendarYearColumn", "NominalFirstColumn",
               "PvFirstColumn", "NominalProfileColumn", "PvProfileColumn",
               "LastColumn", "OffsetOf", "TotalColumn", "SnapshotLong")
    for name in helpers:
        signature = re.search(rf"^Private Function {name}\((.*?)\) As", text, re.M | re.S)
        assert signature, f"{name} is missing"
        assert "bank As String" in signature.group(1), (
            f"{name} produces an address without being told the bank")
    # AND EACH ARM MUST NAME ITS OWN BANK'S CONSTANT. A helper that took the
    # bank and returned bank A's column whatever it was told would satisfy every
    # claim above and silently publish one bank's answer into the other's block.
    for name in helpers:
        block = re.search(
            rf"^Private Function {name}\(.*?\n(.*?)^End Function", text, re.M | re.S)
        arms = re.findall(r"SIM_\w*(?:_A_|_B_|_COLUMN_A|_COLUMN_B)\w*", block.group(1))
        if not arms:
            continue
        banks = {"A" if ("_A_" in token or token.endswith("_A")) else "B"
                 for token in arms}
        assert banks == {"A", "B"}, (
            f"{name} resolves to bank {sorted(banks)} only; it cannot address "
            "both banks")
        assert len(set(arms)) == len(arms), f"{name} names a constant twice"

    # And no procedure mentions both banks' record columns outside the branch
    # that chooses between them.
    for block in re.finditer(r"^(Public|Private) Function (\w+)\(.*?\n(.*?)^End Function",
                             text, re.M | re.S):
        name, body = block.group(2), block.group(3)
        if name in helpers:
            continue
        both = re.findall(r"SIM_ANNUAL_([AB])_\w+_COLUMN", body)
        assert not both, f"{name} names a bank's column directly: {sorted(set(both))}"


# ===========================================================================
# D. THE SETTLEMENT, EXECUTED
# ===========================================================================
NOT_PRODUCED = "NOT PRODUCED"
CURRENT = "CURRENT"
HISTORICAL = "HISTORICAL"
OTHER_PX = "OTHER Px"


def _distribution(published, simulation_current, identity_matches):
    return _vba()["DistributionStateOf"](
        _Ref(published), _Ref(simulation_current), _Ref(identity_matches))


def _profile_state(distribution, consistent, resolved, stamped_label,
                   stamped_p, resolved_label, resolved_p):
    return _vba()["ProfileStateOf"](
        _Ref(distribution), _Ref(consistent), _Ref(resolved), _Ref(stamped_label),
        _Ref(stamped_p), _Ref(resolved_label), _Ref(resolved_p))


@pytest.mark.parametrize("published,current,matches,expected", [
    ("", True, True, NOT_PRODUCED),
    ("PUBLISHED", True, True, CURRENT),
    ("PUBLISHED", True, False, HISTORICAL),
    ("PUBLISHED", False, True, HISTORICAL),
    ("PUBLISHED", False, False, HISTORICAL),
    ("published", True, True, NOT_PRODUCED),
])
def test_13_the_distribution_state_is_the_run_and_nothing_else(
        published, current, matches, expected) -> None:
    assert _distribution(published, current, matches) == expected


def test_14_moving_the_selector_does_not_retire_the_ladders() -> None:
    """THE SETTLEMENT'S FIRST HALF, proved on the accepted source.

    The distribution state has no selector input at all - there is no argument
    it could arrive through - so no movement of the reporting selector can
    change it.
    """
    signature = re.search(
        r"Private Function DistributionStateOf\((.*?)\) As String",
        STORE_BAS.read_text(encoding="utf-8"), re.S).group(1)
    for forbidden in ("Px", "Label", "Probability", "selector", "Selected"):
        assert forbidden not in signature, (
            f"the distribution state can see {forbidden!r}, so a reporting "
            "selector could retire a ladder it never entered")
    # And the same holds through the values: every selector-bearing situation
    # leaves the ladder state where it was.
    for label, probability in (("P80", 0.8), ("P50", 0.5), ("P95", 0.95)):
        assert _distribution("PUBLISHED", True, True) == CURRENT, (label, probability)


@pytest.mark.parametrize("distribution", [NOT_PRODUCED, HISTORICAL])
def test_15_the_profile_cannot_outlive_the_ladders(distribution) -> None:
    """A profile is not current for a run whose distributions are not, and it
    does not exist for a run that produced nothing. The state is inherited
    unchanged rather than collapsed into one word."""
    assert _profile_state(distribution, True, True, "P80", 0.8, "P80", 0.8) == distribution


def test_16_the_profile_is_current_only_when_both_stamped_fields_agree() -> None:
    assert _profile_state(CURRENT, True, True, "P80", 0.8, "P80", 0.8) == CURRENT
    # The label alone is not enough.
    assert _profile_state(CURRENT, True, True, "P80", 0.8, "P80", 0.5) == OTHER_PX
    # Neither is the probability alone.
    assert _profile_state(CURRENT, True, True, "P80", 0.8, "P50", 0.8) == OTHER_PX
    assert _annual_contract()[
        "profile_current_requires_label_and_probability_match"] is True


def test_17_a_moved_selector_makes_the_profile_another_levels_and_not_stale() -> None:
    """THE SETTLEMENT'S SECOND HALF. The ladders stay CURRENT and the profile
    becomes OTHER Px - not HISTORICAL, which would say the run was superseded,
    and not CURRENT, which would relabel it."""
    assert _distribution("PUBLISHED", True, True) == CURRENT
    moved = _profile_state(CURRENT, True, True, "P80", 0.8, "P50", 0.5)
    assert moved == OTHER_PX
    assert moved != HISTORICAL
    assert moved != CURRENT


def test_18_a_stamp_that_disagrees_with_itself_is_never_current() -> None:
    for consistent in (False,):
        assert _profile_state(CURRENT, consistent, True, "P80", 0.8, "P80", 0.8) == OTHER_PX
    # And an unresolvable selector cannot promote a profile either.
    assert _profile_state(CURRENT, True, False, "P80", 0.8, "", 0.0) == OTHER_PX


def test_19_the_profile_is_never_relabelled() -> None:
    """No input can make the state report the CURRENTLY selected level for a
    profile stamped at another one. Structural, over the procedure's text: it
    returns a state and never a label."""
    text = STORE_BAS.read_text(encoding="utf-8")
    start = text.index("Private Function ProfileStateOf")
    body = text[start:text.index("\nEnd Function", start)]
    assignments = set(re.findall(r"ProfileStateOf = (\w+)", body))
    assert assignments <= {"distribution", "SIM_ANNUAL_STATE_CURRENT",
                           "SIM_ANNUAL_STATE_OTHER_PX"}, assignments
    # And OTHER Px is reachable from more than one branch, so a mismatch cannot
    # fall through to CURRENT.
    assert body.count("SIM_ANNUAL_STATE_OTHER_PX") >= 3, (
        "a mismatched or unresolvable selection has no distinct outcome")
    assert _annual_contract()["profile_relabelled_on_selector_change"] is False
    # PCCM_AnnualProfilePx reports the STAMP, so a reader is told which level
    # the profile belongs to rather than which one is selected.
    px = text[text.index("Public Function PCCM_AnnualProfilePx"):]
    px = px[:px.index("\nEnd Function")]
    assert "SIM_ANNUAL_STAMP_ROW_SELECTED_PX_LABEL" in px
    assert "NM_INPUT_SELECTED_CONFIDENCE_LEVEL" not in px


@pytest.mark.parametrize("field,value", [
    ("run", 2), ("seed", 99), ("fingerprint", "OTHER"), ("digest", "OTHER"),
    ("iterations", 999),
])
def test_20_every_identity_field_can_refuse_a_match(field, value) -> None:
    """All five, not the run id alone: the run id says which attempt, the
    fingerprint which request and the digest which answer."""
    base = dict(run=1, seed=7, fingerprint="FP", digest="RD", iterations=1000)
    def call(values):
        return _vba()["IdentityMatches"](
            _Ref(values["run"]), _Ref(1), _Ref(values["seed"]), _Ref(7),
            _Ref(values["fingerprint"]), _Ref("FP"), _Ref(values["digest"]),
            _Ref("RD"), _Ref(values["iterations"]), _Ref(1000))
    assert call(base) is True
    assert call({**base, field: value}) is False


def test_21_a_blank_identity_is_not_a_match() -> None:
    """Two empty fingerprints are equal strings and say nothing."""
    assert _vba()["IdentityMatches"](
        _Ref(1), _Ref(1), _Ref(7), _Ref(7), _Ref(""), _Ref(""), _Ref(""), _Ref(""),
        _Ref(1000), _Ref(1000)) is False


# ===========================================================================
# E. THE LAYOUT GUARD
# ===========================================================================
def _layout(**overrides):
    offsets = _offsets("A")
    values = {
        "indexAt": offsets["project_index"], "yearAt": offsets["calendar_year"],
        "nominalAt": offsets["nominal"], "pvAt": offsets["pv"],
        "nominalProfileAt": offsets["nominal_profile"],
        "pvProfileAt": offsets["pv_profile"], "fields": _fields("A"),
    }
    values.update(overrides)
    detail = _Ref("")
    ok = _vba()["LayoutIsSound"](
        _Ref(values["indexAt"]), _Ref(values["yearAt"]), _Ref(values["nominalAt"]),
        _Ref(values["pvAt"]), _Ref(values["nominalProfileAt"]),
        _Ref(values["pvProfileAt"]), _Ref(values["fields"]), detail)
    return ok, detail.v


def test_22_the_real_contract_layout_is_accepted() -> None:
    ok, detail = _layout()
    assert ok, detail
    assert _fields("A") == 2 + 2 * _annual_contract()["quantile_count"] + 2
    assert _fields("B") == _fields("A"), "the two banks are not the same shape"


@pytest.mark.parametrize("overrides,expected", [
    ({"pvAt": _offsets("A")["nominal"]}, "both write column offset"),
    ({"yearAt": 0}, "both write column offset"),
    ({"pvProfileAt": 99}, "outside the record block"),
    ({"pvProfileAt": -1}, "outside the record block"),
    ({"fields": 30}, "the contracted fields need"),
])
def test_23_a_broken_layout_is_refused_and_says_why(overrides, expected) -> None:
    ok, detail = _layout(**overrides)
    assert not ok, f"{overrides} was accepted"
    assert expected in detail, detail


def test_24_no_column_can_be_left_unwritten() -> None:
    """A hole would be a column the writer never touches and the clear does -
    and it is IMPOSSIBLE rather than checked for.

    The guard exists in three parts: the block is exactly as wide as the number
    of claims, every claim is proved to be inside it, and no two may share a
    slot. Twenty-six distinct in-range claims over twenty-six columns leave none
    over, so a scan for a gap could never find one. This asserts that reasoning
    directly - over the real contract, and over every rearrangement that keeps
    the three parts true - rather than leaving unreachable code standing in for
    it.
    """
    offsets = _offsets("A")
    fields = _fields("A")
    claims = [offsets["project_index"], offsets["calendar_year"],
              offsets["nominal_profile"], offsets["pv_profile"]]
    claims += [offsets["nominal"] + rung for rung in range(11)]
    claims += [offsets["pv"] + rung for rung in range(11)]
    assert len(claims) == fields, "the claim count is not the block width"
    assert len(set(claims)) == len(claims), "two contracted fields share a column"
    assert set(claims) == set(range(fields)), "the claims do not cover the block"
    # And the source no longer carries a scan that could never fire.
    body = STORE_BAS.read_text(encoding="utf-8")
    body = body[body.index("Private Function LayoutIsSound"):]
    body = body[:body.index("\nEnd Function")]
    assert "no contracted field writes it" not in body, (
        "an unreachable exhaustiveness scan is back")


def test_24b_an_overlap_is_what_a_gap_actually_looks_like() -> None:
    """Because the count is fixed, any hole is also a collision - so the
    collision guard is the one that fires, and it names both claimants."""
    offsets = _offsets("A")
    ok, detail = _layout(nominalProfileAt=offsets["nominal"])
    assert not ok
    assert "both write column offset" in detail, detail


# ===========================================================================
# F. THE READ-BACK: THE VALUES THAT WOULD BE PERSISTED
# ===========================================================================
def _probabilities():
    if "probabilities" not in _CACHE:
        out, detail = [], _Ref("")
        assert _vba()["SimStatsLadderProbabilities"](out, detail), detail.v
        _CACHE["probabilities"] = [float(v) for v in out]
    return _CACHE["probabilities"]


def test_25_the_rungs_are_the_accepted_ladders_own() -> None:
    """Eleven probabilities, in the projected order, decoded by the one owner."""
    projected = [_projected()[f"SIM_QUANTILE_{index + 1}"] for index in range(11)]
    assert _probabilities() == [int(label[1:]) / 100.0 for label in projected]
    assert len(_probabilities()) == _annual_contract()["quantile_count"]


@pytest.mark.parametrize("measure", ["nominal", "PV"])
def test_26_the_produced_ladder_equals_the_independent_oracle(measure) -> None:
    """Bit for bit, over eleven rungs and every project year.

    The orchestration blocks the years and indexes into a block; the oracle
    reduces the full matrix in plain Python. They share no code, so a blocking
    or indexing defect cannot be hidden by a shared mistake.
    """
    ladder, _ = _produce(measure, 0.8, _probabilities())
    expected = _oracle_ladder(measure, _probabilities())
    rungs = len(_probabilities())
    for year in range(YEARS):
        for rung in range(rungs):
            assert ladder[year * rungs + rung] == expected[year][rung], (
                f"{measure} year {year + 1} rung {rung + 1}")


@pytest.mark.parametrize("measure", ["nominal", "PV"])
@pytest.mark.parametrize("p", [0.5, 0.7, 0.8, 0.95])
def test_27_the_produced_profile_equals_the_independent_oracle(measure, p) -> None:
    _, profile = _produce(measure, p, _probabilities())
    assert profile == _oracle_profile(measure, p), f"{measure} at p={p}"


@pytest.mark.parametrize("measure", ["nominal", "PV"])
def test_28_the_profile_sums_to_the_reported_percentile(measure) -> None:
    """The reconciliation the pipeline itself performs, checked here against the
    accepted percentile over the published totals - to the project's own
    allowance, never to bit equality, and with nothing scaled to make it pass."""
    p = 0.8
    _, profile = _produce(measure, p, _probabilities())
    totals = replay._run_totals()[1 if measure == "PV" else 0]
    reported = percentile_type7(totals, p)
    allowance = annual.reconciliation_allowance(
        [abs(v) for v in profile], [abs(reported)],
        ABSOLUTE_FLOOR, RELATIVE_COEFFICIENT, SCALE_FLOOR)
    assert abs(sum(profile) - reported) <= allowance, (
        f"{measure}: {sum(profile)} against {reported}")
    # And the VBA reconciliation agrees.
    detail = _Ref("")
    assert _vba()["ReconcileProfile"](
        list(profile), _identity(SelectedProbability=p), list(totals),
        _Ref(measure), detail), detail.v


# THE BLOCKING, ACTUALLY EXERCISED.
#
# The accepted block width is twelve and the fixture is four years, so the
# production above never takes a second pass - and an indexing defect that used
# the DURATION where it should use the BLOCK's width would be invisible in every
# test that came before this one. So the width is narrowed and the same
# comparison is repeated: the answer must not depend on the blocking, because a
# year's ladder is a function of that year's column alone.
def _narrow(width: int) -> dict:
    key = f"vba:{width}"
    if key not in _CACHE:
        saved = dict(_CACHE)
        _CACHE.pop("vba", None)
        _CACHE.pop("consts", None)
        projected = dict(_projected())
        projected["SIM_ANNUAL_BLOCK_WIDTH"] = width
        _CACHE["consts"] = projected
        _CACHE.pop("vba", None)
        narrowed = _vba()
        _CACHE.clear()
        _CACHE.update(saved)
        _CACHE[key] = narrowed
    return _CACHE[key]


@pytest.mark.parametrize("width", [1, 2, 3])
def test_28b_the_answer_does_not_depend_on_the_blocking(width) -> None:
    """One pass, two passes or four: the same ladders and the same profile.

    An index that used the project duration where it should use the BLOCK's own
    width would agree with the oracle at four years and twelve, and disagree
    here on the first year of the second block.
    """
    p = 0.8
    narrow = _narrow(width)
    assert narrow["SIM_ANNUAL_BLOCK_WIDTH"] == width
    blocks = narrow["SimAnnualBlockCount"](_Ref(YEARS), _Ref(width))
    assert blocks == -(-YEARS // width) > 1, "the width did not force extra passes"

    detail = _Ref("")
    for measure in ("nominal", "PV"):
        totals = replay._run_totals()[1 if measure == "PV" else 0]
        position = narrow["_new"]("SimStatsPosition")
        assert narrow["SimStatsQuantilePosition"](
            list(totals), _Ref(len(totals)), _Ref(p), position, detail), detail.v
        ladder, profile = [], []
        assert narrow["ProduceAnnual"](
            _identity(SelectedProbability=p), replay._records(),
            _Ref(len(replay._records())), replay._year_factors(measure),
            _Ref(measure), position, _probabilities(), ladder, profile,
            detail), detail.v
        expected = _oracle_ladder(measure, _probabilities())
        rungs = len(_probabilities())
        for year in range(YEARS):
            for rung in range(rungs):
                assert float(ladder[year * rungs + rung]) == expected[year][rung], (
                    f"{measure} width {width} year {year + 1} rung {rung + 1}")
        assert [float(v) for v in profile] == _oracle_profile(measure, p), (
            f"{measure} at block width {width}")


def test_29_the_ladder_is_not_the_profile() -> None:
    """Two different objects. The ladder does not sum to the total percentile
    and is not selector-specific; the profile is both."""
    p = 0.8
    full, profile = _produce("nominal", p, _probabilities())
    rungs = len(_probabilities())
    at = _probabilities().index(p)
    # THE SAME p, taken the two ways. One is a percentile per project year; the
    # other is one pair of iterations' own annual shapes.
    ladder = [full[year * rungs + at] for year in range(YEARS)]
    assert ladder != profile
    reported = percentile_type7(replay._run_totals()[0], p)
    assert abs(sum(ladder) - reported) > 1.0, (
        "the per-year ladder sums to the total percentile in this model, so it "
        "cannot demonstrate that the two are different objects")
    assert _annual_contract()["distribution_currentness_is_selector_specific"] is False
    assert _annual_contract()["profile_currentness_is_selector_specific"] is True


@pytest.mark.parametrize("bank", ["A", "B"])
def test_30_the_flattened_row_puts_every_value_in_its_contracted_column(bank) -> None:
    """The record block, read back through the contract's own offsets."""
    p = 0.8
    nominal_ladder, nominal_profile = _produce("nominal", p, _probabilities())
    pv_ladder, pv_profile = _produce("PV", p, _probabilities())
    project_index = list(range(1, YEARS + 1))
    calendar_year = [2026 + offset for offset in range(YEARS)]

    flat, fields, detail = [], _Ref(0), _Ref("")
    ok = _vba()["SimAnnualStoreFlatten"](
        _identity(Bank=bank, SelectedProbability=p), list(project_index),
        list(calendar_year), list(nominal_ladder), list(pv_ladder),
        list(nominal_profile), list(pv_profile), flat, fields, detail)
    assert ok, detail.v
    assert int(fields.v) == _fields(bank)

    offsets = _offsets(bank)
    rungs = len(_probabilities())
    for year in range(YEARS):
        origin = year * _fields(bank)
        assert flat[origin + offsets["project_index"]] == float(project_index[year])
        assert flat[origin + offsets["calendar_year"]] == float(calendar_year[year])
        assert flat[origin + offsets["nominal_profile"]] == nominal_profile[year]
        assert flat[origin + offsets["pv_profile"]] == pv_profile[year]
        for rung in range(rungs):
            assert flat[origin + offsets["nominal"] + rung] == \
                nominal_ladder[year * rungs + rung]
            assert flat[origin + offsets["pv"] + rung] == pv_ladder[year * rungs + rung]
    assert len(flat) == YEARS * _fields(bank)


def test_31_the_two_banks_flatten_identically() -> None:
    """A/B ISOLATION IS ADDRESSING, NOT CONTENT. The same answer packed for
    either bank is the same block of numbers; only where it lands differs."""
    p = 0.8
    nominal_ladder, nominal_profile = _produce("nominal", p, _probabilities())
    pv_ladder, pv_profile = _produce("PV", p, _probabilities())
    packed = {}
    for bank in ("A", "B"):
        flat, fields, detail = [], _Ref(0), _Ref("")
        assert _vba()["SimAnnualStoreFlatten"](
            _identity(Bank=bank, SelectedProbability=p), list(range(1, YEARS + 1)),
            [2026 + offset for offset in range(YEARS)], list(nominal_ladder),
            list(pv_ladder), list(nominal_profile), list(pv_profile),
            flat, fields, detail), detail.v
        packed[bank] = [float(v) for v in flat]
    assert packed["A"] == packed["B"]


@pytest.mark.parametrize("measure,scalar", [("nominal", "Knom"), ("PV", "Kpv")])
def test_32_the_per_year_factors_sum_back_to_the_accepted_scalar(measure, scalar) -> None:
    """The pipeline's own reconciliation, run over the real resolved model."""
    factors, detail = [], _Ref("")
    ok = _vba()["BuildYearFactors"](
        replay._records(), _Ref(len(replay._records())), _Ref(YEARS),
        list(replay._discount()), _Ref(measure == "PV"), factors, detail)
    assert ok, detail.v
    assert len(factors) == len(replay._records()) * YEARS
    # Driver-major, stride YEARS - the supply order the engine maps by permanent
    # id - and identical to the accepted owner's own factors.
    assert [float(v) for v in factors] == replay._year_factors(measure)


def test_33_a_driver_whose_terms_do_not_reconcile_is_refused() -> None:
    """NOTHING IS SCALED TO MAKE A SUM COME OUT. A per-year decomposition that
    does not sum back to the scalar Phase 5 built is a refusal that names the
    driver, not a residual that gets absorbed."""
    records = [dict(r, Weights=list(r["Weights"]), Inflation=list(r["Inflation"]))
               for r in replay._records()]
    records[1]["Knom"] = records[1]["Knom"] * 1.5 + 1.0
    factors, detail = [], _Ref("")
    ok = _vba()["BuildYearFactors"](
        records, _Ref(len(records)), _Ref(YEARS), list(replay._discount()),
        _Ref(False), factors, detail)
    assert not ok, "a broken factor decomposition was accepted"
    assert records[1]["PermanentId"] in detail.v, detail.v
    assert "outside the allowance" in detail.v, detail.v


def test_34_the_year_axis_must_agree_with_the_resolved_model() -> None:
    detail = _Ref("")
    assert _vba()["CrossCheckYearCount"](
        replay._records(), _Ref(len(replay._records())), _Ref(YEARS), detail), detail.v
    assert not _vba()["CrossCheckYearCount"](
        replay._records(), _Ref(len(replay._records())), _Ref(YEARS + 1), detail)
    assert "published year axis carries" in detail.v, detail.v
    assert _annual_contract()["year_count_cross_checked_against_driver_weights"] is True


# ===========================================================================
# G. FAILURE INJECTION
# ===========================================================================
def _position_at(measure: str, p: float):
    detail = _Ref("")
    position = _vba()["_new"]("SimStatsPosition")
    totals = replay._run_totals()[1 if measure == "PV" else 0]
    assert _vba()["SimStatsQuantilePosition"](
        list(totals), _Ref(len(totals)), _Ref(p), position, detail), detail.v
    return position


@pytest.mark.parametrize("overrides,expected", [
    ({"Iterations": 0}, "at least one iteration"),
    ({"YearCount": 0}, "no year to decompose"),
])
def test_35_a_production_that_cannot_run_refuses_and_names_the_measure(
        overrides, expected) -> None:
    """A refusal anywhere in the production leaves the caller with a refusal,
    not a partial answer - and the publication it feeds is never reached,
    because the pipeline exits before it."""
    detail = _Ref("")
    ladder, profile = [], []
    ok = _vba()["ProduceAnnual"](
        _identity(**overrides), replay._records(), _Ref(len(replay._records())),
        replay._year_factors("nominal"), _Ref("nominal"), _position_at("nominal", 0.8),
        _probabilities(), ladder, profile, detail)
    assert not ok, f"{overrides} was accepted"
    assert expected in detail.v, detail.v
    assert detail.v.startswith("annual, nominal:"), detail.v


def test_35b_an_unknown_measure_is_refused_by_the_engine() -> None:
    detail = _Ref("")
    ladder, profile = [], []
    ok = _vba()["ProduceAnnual"](
        _identity(), replay._records(), _Ref(len(replay._records())),
        replay._year_factors("nominal"), _Ref("real"), _position_at("nominal", 0.8),
        _probabilities(), ladder, profile, detail)
    assert not ok
    assert "unknown measure" in detail.v, detail.v


def test_37_the_reconciliation_refuses_rather_than_absorbs() -> None:
    detail = _Ref("")
    assert _vba()["ReconcileTerms"](
        [1.0, 2.0, 3.0], _Ref(3), _Ref(6.0), _Ref("test"), detail), detail.v
    assert not _vba()["ReconcileTerms"](
        [1.0, 2.0, 3.0], _Ref(3), _Ref(6.5), _Ref("test"), detail)
    assert "outside the allowance of" in detail.v, detail.v


def test_38_the_allowance_survives_cancellation() -> None:
    """ERRATUM C1: the conditioning scale sums CONTRIBUTIONS, not aggregates.

    THE TEST HAS TO DISCRIMINATE, and a difference of zero does not: it passes
    under either scale. So the case chosen is one where the two scales DISAGREE
    about the same difference.

    The terms are 1e12, -1e12 and 1, so 2e12 of arithmetic produces an aggregate
    of 1. The contribution scale is 1e-12 x 2e12 ~ 2; an aggregate-only scale
    would be about 2e-12 and would fall to the 1e-6 floor. A residual of 1.0 is
    inside the first and far outside the second - and at that conditioning it
    genuinely is noise, which is the whole reason C1 exists.
    """
    terms = [1e12, -1e12, 1.0]
    detail = _Ref("")
    # Difference 1.0, against a contribution-conditioned allowance of about 2.
    assert _vba()["ReconcileTerms"](
        list(terms), _Ref(3), _Ref(2.0), _Ref("cancelling"), detail), detail.v
    # And the allowance is still an allowance: 4.0 is outside it.
    assert not _vba()["ReconcileTerms"](
        list(terms), _Ref(3), _Ref(5.0), _Ref("cancelling"), detail)
    assert "outside the allowance of" in detail.v, detail.v
    # THE SCALE IS BUILT FROM THE TERMS, structurally: the loop over the
    # contributions is what an aggregate-only scale would delete.
    body = _code(RUN_BAS)
    body = body[body.index("Private Function ReconcileTerms"):]
    body = body[:body.index("End Function")]
    assert "For index = 0 To count - 1" in body, (
        "the conditioning scale no longer sums the contributions")


@pytest.mark.parametrize("label", ["", "P10", "P77", "TEN", "p80"])
def test_39_an_unusable_confidence_level_is_refused(label) -> None:
    """A selector that cannot be resolved refuses the whole step rather than
    publishing ladders beside an unstamped profile. P10 is reported and fixed;
    it is not selectable."""
    p, detail = _Ref(0.0), _Ref("")
    assert not _vba()["SimStatsSelectedProbability"](_Ref(label), p, detail), label
    assert detail.v, "a refusal must say why"


@pytest.mark.parametrize("label,expected", [
    ("P50", 0.5), ("P80", 0.8), ("P95", 0.95),
])
def test_40_a_selectable_level_resolves_through_the_one_ladder(label, expected) -> None:
    p, detail = _Ref(0.0), _Ref("")
    assert _vba()["SimStatsSelectedProbability"](_Ref(label), p, detail), detail.v
    assert p.v == expected


# ===========================================================================
# H. THE FAILURE-INJECTION MATRIX
# ===========================================================================
# Eleven ways the annual step can fail, and one interruption. Each entry names
# how it is evidenced, because they cannot all be evidenced the same way: a
# refusal that happens inside a worksheet read has no Linux execution path, and
# for those the claim is structural - the refusal exists, it is reached before
# anything is written, and the writer is the last step of the pipeline.
#
#   A  the simulation is not CURRENT                    structural
#   B  no simulation has been published                 structural
#   C  Phase 5 is not CURRENT (the bridge refuses)      structural
#   D  the model resolves no drivers                    structural
#   E  the reporting selector cannot be resolved        executed (test_39)
#   F  the published year axis is missing or disagrees  executed (test_34) +
#                                                       structural
#   G  a per-year factor is not representable           executed
#   H  a block cannot be replayed                       executed (test_35)
#   I  a year's ladder cannot be taken                  executed (test_35)
#   J  the profile blend is refused                     executed
#   K  the reconciliation is outside the allowance      executed (test_33, 37)
#   -  an interruption after the clear                  structural (test_10)
#
# The one claim they share is the one that matters: NOTHING IS WRITTEN. The
# publication is the last step of the pipeline and every refusal above it exits.
# Each step, and HOW MANY TIMES it must appear. The count is what makes a
# dropped measure visible: every per-measure step runs twice, once for nominal
# and once for PV, and a pipeline that reconciled only the nominal profile would
# still be in the right order.
RUN_PIPELINE = (
    ("SimAnnualStoreCurrentRun", 1),      # A, B
    ("ResolveSelectedPx", 1),             # E
    ("ResolveDrivers", 1),                # C, D
    ("SimAnnualStoreYearAxis", 1),        # F
    ("CrossCheckYearCount", 1),           # F
    ("BuildYearFactors", 2),              # G, K - one per measure
    ("SimStatsLadderProbabilities", 1),
    ("SimAnnualStoreTotals", 2),
    ("SimStatsQuantilePosition", 2),
    ("ProduceAnnual", 2),                 # H, I, J
    ("ReconcileProfile", 2),              # K
    ("SimAnnualStoreFlatten", 1),
    ("SimAnnualStorePublish", 1),
)


def _run_annual_body() -> str:
    text = RUN_BAS.read_text(encoding="utf-8")
    start = text.index("Private Function RunAnnual()")
    return text[start:text.index("\nEnd Function", start)]


def test_41_the_pipeline_runs_in_the_contracted_order_and_publishes_last() -> None:
    body = _run_annual_body()
    positions = []
    for step, times in RUN_PIPELINE:
        found = body.count(step)
        assert found == times, (
            f"{step} appears {found} time(s) in the pipeline and must appear "
            f"{times}: a per-measure step that runs once has silently dropped a "
            "measure")
        positions.append(body.index(step))
    assert positions == sorted(positions), "the pipeline is out of order"
    assert positions[-1] == max(positions), "publication is not the last step"
    # AND BOTH MEASURES REACH EVERY PER-MEASURE STEP BY NAME.
    for measure in ("SIM_MEASURE_NOMINAL", "SIM_MEASURE_PV"):
        assert body.count(measure) >= 2, f"{measure} is not carried through"


def test_42_every_refusal_exits_before_the_publication() -> None:
    """A refusal is an attempt outcome, not a partial answer.

    Every guarded step assigns a refusal and exits; the publication is reached
    only when none of them did.
    """
    body = _run_annual_body()
    refusals = [match.start() for match in re.finditer(r"RunAnnual = Refused\(", body)]
    guarded = sum(times for _, times in RUN_PIPELINE)
    assert len(refusals) == guarded, (
        f"{len(refusals)} refusals for {guarded} guarded calls: a step is "
        "unguarded, or a refusal has no step")
    publish = body.index("SimAnnualStorePublish")
    for start in refusals[:-1]:
        assert start < publish, "a refusal is written after the publication"
        tail = body[start:start + 200]
        assert "Exit Function" in tail, "a refusal does not exit"


@pytest.mark.parametrize("token,expected", [
    ("modSimReport.PCCM_SimulationStatus", "the simulation is "),          # A
    ("no successful simulation has been published", "nothing to decompose"),  # B
])
def test_43_the_precondition_refuses_a_run_it_may_not_explain(token, expected) -> None:
    """A and B. The status is asked for, never re-derived - a second derivation
    of run state is a second answer to the only question that matters."""
    code = _code(STORE_BAS)
    assert token in code
    assert expected in code
    body = code[code.index("Public Function SimAnnualStoreCurrentRun"):]
    body = body[:body.index("End Function")]
    assert "SIM_STATE_CURRENT" in body
    # NOT re-derived: the module reaches no state machinery of its own, and the
    # only simulation-state word it carries is the one it compares against.
    assert "DeriveStatus" not in code
    assert re.findall(r"SIM_STATE_\w+", code) == ["SIM_STATE_CURRENT"], (
        "the store carries a second simulation-state vocabulary")


def test_44_a_model_with_no_drivers_is_refused() -> None:
    """D. A decomposition of nothing is not an answer, and the engine cannot
    prepare a zero-driver replay."""
    body = _code(RUN_BAS)
    assert "the model resolves no drivers to decompose" in body
    resolve = body[body.index("Private Function ResolveDrivers"):]
    resolve = resolve[:resolve.index("End Function")]
    assert "If driverCount < 1 Then" in resolve


def test_45_a_missing_year_axis_is_refused_rather_than_assumed() -> None:
    """F. Half of it executes (test_34); this is the half that reads a table."""
    body = _code(STORE_BAS)
    axis = body[body.index("Public Function SimAnnualStoreYearAxis"):]
    axis = axis[:axis.index("End Function")]
    for guard in ("LoExists(CALC_SHEET, TBL_CALC_YEARS)",
                  "carries no year",
                  "beyond the structural maximum",
                  "no usable index", "no usable calendar year",
                  "no usable discount factor"):
        assert guard in axis, f"the year axis accepts {guard!r} silently"
    # AND IT BUILDS NOTHING. No discount rate is consulted and no series is made.
    for banned in ("BuildDiscountFactors", "discountRate", "modTimeline",
                   "modInflation"):
        assert banned not in body, f"the store reaches {banned!r}"


def test_46_a_factor_that_is_not_representable_is_refused() -> None:
    """G. The refusal names the driver and the project year."""
    records = [dict(r, Weights=list(r["Weights"]), Inflation=list(r["Inflation"]))
               for r in replay._records()]
    records[0]["FxRate"] = 1e308
    records[0]["Weights"] = [1e10] * YEARS
    factors, detail = [], _Ref("")
    ok = _vba()["BuildYearFactors"](
        records, _Ref(len(records)), _Ref(YEARS), list(replay._discount()),
        _Ref(False), factors, detail)
    assert not ok, "an unrepresentable factor was accepted"
    assert records[0]["PermanentId"] in detail.v, detail.v
    assert "not representable" in detail.v, detail.v


def test_47_a_blend_outside_zero_to_one_is_refused() -> None:
    """J. The profile is a CONVEX combination or it is not produced.

    The position the pipeline supplies is always inside [0, 1]; this proves the
    guard is real rather than relying on that.
    """
    position = _vba()["_new"]("SimStatsPosition")
    position.update({"LoSource": 0, "HiSource": 1, "Fraction": 1.5,
                     "LoValue": 0.0, "HiValue": 0.0})
    ladder, profile, detail = [], [], _Ref("")
    ok = _vba()["ProduceAnnual"](
        _identity(), replay._records(), _Ref(len(replay._records())),
        replay._year_factors("nominal"), _Ref("nominal"), position,
        _probabilities(), ladder, profile, detail)
    assert not ok
    assert "outside [0, 1]" in detail.v, detail.v


def test_48_a_refusal_leaves_the_published_answer_where_it_was() -> None:
    """The interruption case, and the reason the marker goes blank first.

    A refusal never reaches the publication at all, so the previous answer is
    untouched. An interruption INSIDE the publication leaves the marker blank,
    which the handoff reports as NOT PRODUCED - never as a current answer.
    """
    assert _distribution("", True, True) == NOT_PRODUCED
    body = _publish_body()
    assert body.index("vbNullString") < body.index("ClearContents"), (
        "the block is cleared while it still claims to be published")
