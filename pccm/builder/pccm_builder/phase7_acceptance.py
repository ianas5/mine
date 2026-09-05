#!/usr/bin/env python3
"""The Phase-7 Windows acceptance projection and its expected-value corpus.

WHY TWO ARTEFACTS AND WHY NEW ONES
----------------------------------
`phase7_acceptance_inspection.json`  WHERE to look. The annual block geometry,
                                     the annual stamp, the handoff vocabulary
                                     and the Phase-7 command surface, projected
                                     from the accepted contracts.
`phase7_acceptance_cases.json`       WHAT to expect. The acceptance fixture
                                     MODELS, and - for the two inherited
                                     Phase-5 scenarios - every published value
                                     the ACCEPTED PHASE-5 ORACLE produces from
                                     them.

NEITHER EXTENDS A PHASE-6 ARTEFACT. `phase6_gate_b_inspection.json` is the
projection Run 6 was accepted against; widening its schema to carry Phase-7
geometry would change an artefact whose identity is historical evidence, and
would make one file answer for two phases' acceptance. The Phase-6 files stay
exactly as Run 6 left them and the Windows harness reads both.

THE MODEL AND ITS EXPECTATION ARE ONE ARTEFACT, DELIBERATELY. If the Windows
harness built its own fixture and this file computed an expectation from a
second copy of the same model, the two would be free to drift and the comparison
would quietly stop meaning anything. So the model is written ONCE, here, in the
shape the accepted `Set-Phase5Fixture` already consumes; PowerShell reads it
rather than constructing it, and the expectation beside it is the accepted
oracle's answer for that exact payload.

THE EXPECTATION IS PRE-PHASE-7 AUTHORITY. `calc_oracle.calculate` is the
independent Phase-5 implementation the accepted `phase5_cases.json` corpus is
built from. Phase 7 changed no line of it. Comparing the live Phase-7 workbook
against it is therefore a regression check against an authority that predates
the change - not a Phase-7 recalculation compared with itself.

IDENTITIES ONLY IN THE INSPECTION. Sheets, columns, rows, cells and procedure
names. No expected number, no tolerance and no vocabulary that is model
semantics - the same line the Phase-5 and Phase-6 projections hold. The
handoff STATE WORDS are the exception and are carried on purpose: they are
contract-owned identifiers a reader must match exactly, exactly as
`published` is, and they are projected from the contract rather than restated.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .artifact_io import write_lf_artifact
from .calc_cases import evaluate, tolerances_from
from .calc_loader import CalcContract
from .sim_loader import SimContract

SCHEMA_VERSION = 1

INSPECTION_FILENAME = "phase7_acceptance_inspection.json"
CASES_FILENAME = "phase7_acceptance_cases.json"

# THE POSITIVE SCHEMA, on the terms the Phase-6 projection established: an
# allowlist refuses the next unforeseen key, a banned list only refuses the ones
# somebody already thought of.
ALLOWED_INSPECTION_KEYS = ("schema_version", "purpose", "provenance",
                           "annual_records", "handoff", "command_surface")
ALLOWED_ANNUAL_KEYS = ("sheet", "header_row", "first_record_row", "quantile_count",
                       "max_record_rows", "index_columns", "quantile_first_column",
                       "selected_px_profile_columns", "stamp")
ALLOWED_STAMP_KEYS = ("bank_value_columns", "rows", "published_marker")
ALLOWED_HANDOFF_KEYS = ("accessors", "distribution_states", "profile_states",
                        "inconsistent_stamp_state")
ALLOWED_CASE_KEYS = ("schema_version", "purpose", "provenance", "scenarios")
ALLOWED_SCENARIO_KEYS = ("id", "title", "purpose", "dimension", "model",
                         "expected", "iterations", "seed_mode", "supplied_seed",
                         "selected_confidence_level", "second_confidence_level",
                         "shrink_model")

# The published marker text. It is production's own constant, projected here
# from the same place `sim_emit` projects it into VBA.
PUBLISHED_MARKER = "PUBLISHED"


# ===========================================================================
# THE PROJECTION
# ===========================================================================
def build_phase7_inspection(sim: SimContract, max_record_rows: int) -> dict[str, Any]:
    raw = sim.raw["sim_data"]
    annual = raw["annual_records"]
    stamp = annual["stamp"]
    handoff = annual["handoff"]
    return {
        "schema_version": SCHEMA_VERSION,
        "purpose": (
            "Where the Phase-7 Windows acceptance harness looks. Addresses and "
            "identities projected from sim_contract.yaml; no expected value, no "
            "tolerance and no bound."
        ),
        "provenance": {
            "sim_contract": "sim_contract.yaml",
            "sim_contract_version": str(sim.version),
        },
        "annual_records": {
            "sheet": str(raw["sheet"]),
            "header_row": int(annual["header_row"]),
            "first_record_row": int(annual["first_record_row"]),
            "quantile_count": int(annual["quantile_count"]),
            # HOW FAR THE CLEAR REACHES. The structural maximum on generated
            # project-year columns is what bounds the record block, and a
            # duration-shrink check has to read past the new answer to prove the
            # old one is gone - so it needs this number, not the year count.
            "max_record_rows": int(max_record_rows),
            "index_columns": {
                bank: {key: str(letter) for key, letter in sorted(columns.items())}
                for bank, columns in sorted(annual["index_columns"].items())
            },
            "quantile_first_column": {
                bank: {measure: str(letter) for measure, letter in sorted(measures.items())}
                for bank, measures in sorted(annual["quantile_first_column"].items())
            },
            "selected_px_profile_columns": {
                bank: {measure: str(letter) for measure, letter in sorted(measures.items())}
                for bank, measures in sorted(annual["selected_px_profile_columns"].items())
            },
            "stamp": {
                "bank_value_columns": {
                    bank: str(column)
                    for bank, column in sorted(stamp["bank_value_columns"].items())
                },
                "rows": {str(field["key"]): int(field["row"])
                         for field in stamp["fields"]},
                "published_marker": PUBLISHED_MARKER,
            },
        },
        "handoff": {
            "accessors": [str(entry["name"]) for entry in handoff["accessors"]],
            "distribution_states": [str(state) for state in handoff["distribution_states"]],
            "profile_states": [str(state) for state in handoff["profile_states"]],
            "inconsistent_stamp_state": str(handoff["inconsistent_stamp_state"]),
        },
        "command_surface": {
            "annual_endpoint": "PCCM_RunAnnualStochastic",
            "handoff_accessors": [str(entry["name"]) for entry in handoff["accessors"]],
        },
    }


# ===========================================================================
# THE ACCEPTANCE FIXTURES
# ===========================================================================
# TWO DIMENSIONS, PROVED SEPARATELY AND ON PURPOSE.
#
# The Phase-7 change to the accepted Phase-5 path is that `DriverFactors` now
# carries two dynamic arrays of length Y per driver. Its two runtime risks are
# INDEPENDENT: many UDT instances each holding arrays, and long arrays inside a
# UDT. Maximising both at once would prove neither more strongly than the two
# smaller cases do together, and would cost far more Windows time.
#
#   W2  many instances     ~300 drivers,  5 years
#   W3  long arrays         10 drivers, 200 years
#
# 200 is the Architecture Lock Revision B maximum on generated project-year
# columns, so W3 is at the structural ceiling of the array length.
_FAMILIES = ("Triangular", "Beta-PERT", "Uniform")
_PROFILES = ("Standard", "Escalated")
_CURRENCIES = ("SAR", "USD")


def _weights(duration: int) -> list[float]:
    """A profile that sums to exactly 1 and puts money in more than one year.

    THE LAST WEIGHT ABSORBS THE REMAINDER, so the sum is exact rather than
    nearly exact: the profiling tolerance is not a licence to hand the workbook
    a fixture that does not add up.
    """
    if duration < 1:
        raise ValueError("a profile needs at least one project year")
    if duration == 1:
        return [1.0]
    share = round(1.0 / duration, 4)
    weights = [share] * (duration - 1)
    weights.append(round(1.0 - share * (duration - 1), 4))
    return weights


def _driver(index: int, is_risk: bool, duration: int) -> dict[str, Any]:
    base = 100.0 + 7.0 * index
    entry = {
        "permanent_id": ("R-%03d" if is_risk else "CL-%03d") % index,
        "distribution": _FAMILIES[(index - 1) % len(_FAMILIES)],
        "currency": _CURRENCIES[(index - 1) % len(_CURRENCIES)],
        "inflation_profile": _PROFILES[(index - 1) % len(_PROFILES)],
        "min_value": base,
        "most_likely": round(base * 1.35, 6),
        "max_value": round(base * 2.10, 6),
        "profile_weights": _weights(duration),
    }
    if is_risk:
        # Never 0 and never 1: a Risk that always or never occurs carries no
        # occurrence variance and exercises the degenerate arm instead.
        entry["probability"] = ((index - 1) % 7 + 1) / 10.0
    else:
        entry["quantity"] = float(1 + (index - 1) % 5)
    return entry


def _inflation(base_year: int, duration: int, start_year: int) -> dict[str, dict[str, Any]]:
    """A rate for every calendar year the span can require.

    The span runs from base_year + 1 to the last project year, so a 200-year
    project needs 200 rates per profile - which is exactly the array length W3
    exists to exercise.
    """
    last = start_year + duration - 1
    years = range(base_year + 1, last + 1)
    return {
        "Standard": {str(year): 0.03 for year in years},
        "Escalated": {str(year): 0.06 if (year % 2) else 0.055 for year in years},
    }


def _model(driver_count: int, duration: int, base_year: int = 2026,
           start_year: int = 2026) -> dict[str, Any]:
    if driver_count < 2:
        raise ValueError("an acceptance model needs at least two drivers")
    cost_count = max(1, (driver_count * 3) // 5)
    risk_count = driver_count - cost_count
    if risk_count < 1:
        raise ValueError("an acceptance model must carry Risks as well as Cost Lines")
    return {
        "timeline": {"base_year": base_year, "start_year": start_year,
                     "duration": duration},
        "discount_rate": 0.05,
        "fx": [{"currency": "SAR", "rate": 1.0}, {"currency": "USD", "rate": 3.75}],
        "inflation": _inflation(base_year, duration, start_year),
        "cost_lines": [_driver(index, False, duration)
                       for index in range(1, cost_count + 1)],
        "risks": [_driver(index, True, duration)
                  for index in range(1, risk_count + 1)],
    }


def build_phase7_cases(calc: CalcContract, sim: SimContract,
                       max_record_rows: int) -> dict[str, Any]:
    tolerances = tolerances_from(calc)
    w2 = _model(driver_count=300, duration=5)
    w3 = _model(driver_count=10, duration=200)
    # The behavioural fixture: small, deterministic, and shared by W4, W5 and
    # W6 so the annual checks all bind to one identity.
    w4 = _model(driver_count=5, duration=4)
    # W7 reuses one bank after a shorter run, so it needs two durations.
    w7_long = _model(driver_count=5, duration=20)
    w7_short = _model(driver_count=5, duration=4)

    if int(w3["timeline"]["duration"]) != int(max_record_rows):
        raise ValueError(
            "W3 must sit at the structural maximum on generated project-year "
            f"columns: {max_record_rows}, not {w3['timeline']['duration']}")
    if len(w2["cost_lines"]) + len(w2["risks"]) <= len(w3["cost_lines"]) + len(w3["risks"]):
        raise ValueError("W2 must carry more drivers than W3, or it proves nothing W3 does not")
    if int(w2["timeline"]["duration"]) >= int(w3["timeline"]["duration"]):
        raise ValueError("W3 must carry more project years than W2")

    scenarios = [
        {
            "id": "W2", "title": "many DriverFactors instances",
            "dimension": "driver_count",
            "purpose": (
                "Many DriverFactors UDT instances, each carrying two dynamic "
                "arrays. Five project years, so the array LENGTH is not what is "
                "being exercised here."
            ),
            "model": w2, "expected": evaluate(w2, tolerances),
        },
        {
            "id": "W3", "title": "maximum year-array length",
            "dimension": "year_count",
            "purpose": (
                "Dynamic arrays inside DriverFactors at the structural maximum "
                "project-year length. Ten drivers, so the instance COUNT is not "
                "what is being exercised here."
            ),
            "model": w3, "expected": evaluate(w3, tolerances),
        },
        {
            "id": "W4", "title": "base simulation and the no-run refusal",
            "dimension": "behavioural",
            "purpose": (
                "The annual endpoint refused before any simulation exists, then "
                "one deterministic FIXED-seed run the annual scenarios bind to."
            ),
            "model": w4, "expected": evaluate(w4, tolerances),
            "iterations": 1000, "seed_mode": "FIXED", "supplied_seed": 20260905,
            "selected_confidence_level": "P80", "second_confidence_level": "P50",
        },
        {
            "id": "W7", "title": "bank alternation and duration shrink",
            "dimension": "behavioural",
            "purpose": (
                "A to B and back to A, with the reused bank publishing a SHORTER "
                "annual answer than it held before."
            ),
            "model": w7_long, "expected": evaluate(w7_long, tolerances),
            "shrink_model": w7_short,
            "iterations": 1000, "seed_mode": "FIXED", "supplied_seed": 20260906,
            "selected_confidence_level": "P80",
        },
    ]
    return {
        "schema_version": SCHEMA_VERSION,
        "purpose": (
            "The Phase-7 Windows acceptance fixtures, and - for the inherited "
            "Phase-5 scenarios - every published value the ACCEPTED Phase-5 "
            "oracle produces from them. The model is written once and read by "
            "the harness, so the fixture and its expectation cannot drift."
        ),
        "provenance": {
            "calc_contract_version": str(calc.version),
            "sim_contract_version": str(sim.version),
            "expectation_source": "pccm_builder.calc_oracle.calculate",
            # THE COMPARISON ALLOWANCE IS THE PROJECT'S OWN, projected here so
            # the Windows harness cannot invent a broader one. It is the accepted
            # Phase-5 identity ABSOLUTE FLOOR: the smallest difference the
            # project already treats as noise, applied to values many orders of
            # magnitude above it.
            "comparison_absolute_floor": float(calc.tolerances.identity_absolute_floor),
            "expectation_authority": (
                "The independent Phase-5 implementation the accepted "
                "phase5_cases.json corpus is built from. Phase 7 changed no line "
                "of it, so it is pre-Phase-7 authority rather than a Phase-7 "
                "recalculation compared with itself."
            ),
        },
        "scenarios": scenarios,
    }


def _check_keys(node: dict[str, Any], allowed: tuple[str, ...], where: str) -> None:
    extra = sorted(set(node) - set(allowed))
    if extra:
        raise ValueError(f"{where}: unexpected key(s) {extra}")


def validate_phase7_artifacts(inspection: dict[str, Any], cases: dict[str, Any]) -> None:
    _check_keys(inspection, ALLOWED_INSPECTION_KEYS, INSPECTION_FILENAME)
    _check_keys(inspection["annual_records"], ALLOWED_ANNUAL_KEYS,
                f"{INSPECTION_FILENAME}: annual_records")
    _check_keys(inspection["annual_records"]["stamp"], ALLOWED_STAMP_KEYS,
                f"{INSPECTION_FILENAME}: annual_records.stamp")
    _check_keys(inspection["handoff"], ALLOWED_HANDOFF_KEYS,
                f"{INSPECTION_FILENAME}: handoff")
    _check_keys(cases, ALLOWED_CASE_KEYS, CASES_FILENAME)
    for scenario in cases["scenarios"]:
        _check_keys(scenario, ALLOWED_SCENARIO_KEYS,
                    f"{CASES_FILENAME}: scenario {scenario.get('id')}")
    # NO EXPECTED VALUE MAY REACH THE PROJECTION. The two artefacts answer two
    # different questions, and a number in the address file would be a second
    # authority for it. The check covers the DATA sections only - `purpose` and
    # `provenance` are prose about the file, and a rule that could not say so
    # without tripping over its own explanation would be a rule nobody could
    # document.
    payload = {key: value for key, value in inspection.items()
               if key in ("annual_records", "handoff", "command_surface")}
    text = json.dumps(payload)
    # `iterations`, `effective_seed` and `year_count` DO appear in it and are
    # addresses: they are the names of stamp ROWS, which is where to look, not
    # what to expect there. What may never appear is an authority for a VALUE.
    for forbidden in ("tolerance", "expected", "allowance", "budget"):
        if forbidden in text:
            raise ValueError(f"{INSPECTION_FILENAME} carries {forbidden!r}")


def emit_phase7_acceptance(sim: SimContract, calc: CalcContract,
                           max_record_rows: int, build_dir: Path) -> tuple[Path, Path]:
    inspection = build_phase7_inspection(sim, max_record_rows)
    cases = build_phase7_cases(calc, sim, max_record_rows)
    validate_phase7_artifacts(inspection, cases)
    inspection_path = build_dir / INSPECTION_FILENAME
    cases_path = build_dir / CASES_FILENAME
    write_lf_artifact(inspection_path, json.dumps(inspection, indent=2) + "\n")
    write_lf_artifact(cases_path, json.dumps(cases, indent=2) + "\n")
    return inspection_path, cases_path
