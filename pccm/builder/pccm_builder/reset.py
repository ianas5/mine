"""P10-2B: the Reset Results projection.

WHAT THIS EMITS, AND WHY IT IS NOT A SECOND OPINION ABOUT ANYTHING.

A later Windows run has to prove two things about Reset Results that no static
control can prove on its own: that every published result is GONE from the store
that holds it, and that every input and every identity value is EXACTLY what it
was. Both need a list of addresses, and a list of addresses typed by hand would
be a second authority for the layout - the one thing this phase has refused
everywhere else.

So both lists are DERIVED:

    the preserved set    from `resolve_unlocked`, which is the protection
                         owner's own resolution of "every cell a user may type
                         in", plus the applied-timeline INPUTS and the permanent
                         identity counters the structure contract declares;

    the publication set  from the calculation contract's own scalar blocks and
                         table names, and from the simulation contract's own
                         `sim_data` layout.

NO VOLATILE CELL IS IN THE PRESERVED SET. The four derived applied-timeline
cells are formulas, the structural-state cell is a formula, and every status
timestamp is a recalculation artefact. A run that compared those before and
after a reset would fail on a workbook that was behaving correctly, which is
worse than not checking at all - it would make the real check untrustworthy.

THIS FILE RUNS NO RESET AND ASSERTS NOTHING ABOUT ONE. It is a list of places to
look.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .protection import resolve_unlocked
from .sim_loader import (
    LOCKED_ITERATION_BANKS,
    LOCKED_PENDING_AUTO_NONCE_CELL,
    LOCKED_RUN_IDENTITY,
    LOCKED_RUN_IDENTITY_BANK_COLUMNS,
    LOCKED_RUN_IDENTITY_COLUMNS,
    LOCKED_SENSITIVITY_STAMP,
    LOCKED_SENSITIVITY_STAMP_COLUMNS,
)

__all__ = ["resolve_preserved", "resolve_publications", "emit_reset_projection"]

_IDENTITY_ROW = {key: row for key, row, *_rest in LOCKED_RUN_IDENTITY}


def _span(column: str, first: int, last: int) -> str:
    return f"{column}{first}:{column}{last}"


def resolve_preserved(structure: Any, contract: Any, drivers: Any) -> dict[str, Any]:
    """Every value Reset Results must leave exactly as it found it."""
    # THE PROTECTION OWNER'S OWN ANSWER TO "WHAT MAY A USER TYPE IN". Reset must
    # not change any of it, and asking the resolver rather than restating it is
    # what keeps the two answers from ever disagreeing.
    editable = resolve_unlocked(structure, contract, drivers)

    # THE APPLIED TIMELINE, INPUTS ONLY. `structure.applied` is the three stored
    # values; `structure.derived` is the four formulas that follow from them, and
    # a formula is not a preserved value - it is a recomputation.
    applied = [field.cell for field in structure.applied]

    # THE PERMANENT-ID COUNTERS, from the block that declares them.
    counters = [counter["cell"] for counter in structure.identity_block["counters"]]

    return {
        "editable_inputs": [
            {"sheet": sheet, "cells": sorted(addresses)}
            for sheet, addresses in sorted(editable.items()) if addresses
        ],
        "applied_timeline": {"sheet": structure.setup_sheet, "cells": applied},
        "permanent_id_counters": {"sheet": structure.identity_sheet, "cells": counters},
    }


def resolve_publications(calc: Any, sim: Any) -> dict[str, Any]:
    """Every store a completed reset must leave carrying no published result,
    and the simulation identity cells it must leave untouched beside them.

    THE TWO SETS SHARE A COLUMN AND THAT IS THE POINT. `_SimData` column D holds
    the bank-A publication above the AUTO nonce counter, the run-id counter and
    the attempt record; a reset that cleared the whole column would destroy the
    anti-replay guarantee, and a projection that did not say so could not catch
    it. So the rows that must empty and the rows that must not are emitted
    together, from the same contract.
    """
    state = calc.scalar_blocks["calc_state"]
    totals = calc.scalar_blocks["calc_totals"]
    calculation = {
        "sheet": calc.sheet,
        "cleared": {
            "state": _span(state.value_column, state.first_row, state.last_row),
            "totals": _span(totals.value_column, totals.first_row, totals.last_row),
            "tables": list(calc.table_names),
        },
        "attempt_result_initial": {
            "cell": f"{state.value_column}"
                    f"{next(f.row for f in state.fields if f.key == 'last_attempt_result')}",
            # THE CONTRACT'S OWN AS-BUILT VALUE. Reset restores this field to it,
            # which is not a state word chosen by the command: it is what a
            # workbook that has never been calculated carries.
            "value": next(f.initial for f in state.fields
                          if f.key == "last_attempt_result"),
        },
    }

    data = sim.raw["sim_data"]
    value_column = LOCKED_RUN_IDENTITY_COLUMNS["value_column"]
    first_record_row = sim.layout.first_iteration_row

    annual = data["annual_records"]
    annual_stamp_rows = [field["row"] for field in annual["stamp"]["fields"]]
    sensitivity_stamp_rows = [row for _key, row, _type in LOCKED_SENSITIVITY_STAMP]
    # EACH BLOCK'S OWN FIRST AND LAST ROW. The summary ladder and the
    # contingency ladder both start at row 8 and they do NOT end at the same
    # row, so deriving one from the other's rungs would have emitted a
    # contingency range three rows too low - a projection that looked plausible
    # and pointed at the wrong cells.
    summary = data["summary_statistics"]
    contingency = data["contingency_ladder"]

    def annual_last_column(bank: str) -> str:
        # THE BLOCK'S LAST COLUMN IS THE PV PROFILE'S, which is the same
        # derivation modSimAnnualStore.LastColumn makes, from the same contract.
        return annual["selected_px_profile_columns"][bank]["pv"]

    simulation = {
        "sheet": sim.layout.sheet,
        "cleared": {
            "bank_snapshots": {
                bank: _span(column, _IDENTITY_ROW["last_successful_stamp"],
                            _IDENTITY_ROW["applied_timeline"])
                for bank, column in sorted(LOCKED_RUN_IDENTITY_BANK_COLUMNS.items())
            },
            "attempt_and_selector": _span(
                value_column, _IDENTITY_ROW["last_attempt_result"],
                _IDENTITY_ROW["active_bank"]),
            "summary": {
                bank: f"{columns['nominal']}{summary['first_row']}:"
                      f"{columns['pv']}{summary['last_row']}"
                for bank, columns in sorted(summary["bank_value_columns"].items())
            },
            "contingency": {
                bank: f"{columns['nominal']}{contingency['first_row']}:"
                      f"{columns['pv']}{contingency['last_row']}"
                for bank, columns in sorted(contingency["bank_value_columns"].items())
            },
            "iteration_banks": {
                bank: {"columns": [columns[key] for key in
                                   ("iteration_index", "total_nominal", "total_pv")],
                       "first_row": first_record_row}
                for bank, columns in sorted(LOCKED_ITERATION_BANKS.items())
            },
            "annual_stamps": {
                bank: _span(column, min(annual_stamp_rows), max(annual_stamp_rows))
                for bank, column in sorted(annual["stamp"]["bank_value_columns"].items())
            },
            "annual_blocks": {
                bank: {"first_column": annual["index_columns"][bank]["project_index"],
                       "last_column": annual_last_column(bank),
                       "first_row": annual["first_record_row"]}
                for bank in sorted(annual["index_columns"])
            },
            "sensitivity_stamps": {
                bank: _span(column, min(sensitivity_stamp_rows),
                            max(sensitivity_stamp_rows))
                for bank, column in sorted(LOCKED_SENSITIVITY_STAMP_COLUMNS.items())
            },
            "sensitivity_records": {
                bank: {"first_column": span["first_column"],
                       "last_column": span["last_column"],
                       "first_row": data["sensitivity_records"]["first_record_row"]}
                for bank, span in sorted(data["sensitivity_records"]["banks"].items())
            },
        },
        "preserved": {
            "next_auto_nonce": f"{value_column}{_IDENTITY_ROW['next_auto_nonce']}",
            "last_run_id": f"{value_column}{_IDENTITY_ROW['last_run_id']}",
            "pending_auto_nonce": LOCKED_PENDING_AUTO_NONCE_CELL,
        },
    }
    return {"calculation": calculation, "simulation": simulation}


def emit_reset_projection(path: Path, structure: Any, contract: Any, drivers: Any,
                          calc: Any, sim: Any) -> dict[str, Any]:
    projection = {
        "command": "PCCM_ResetResults",
        "preserved": resolve_preserved(structure, contract, drivers),
        "publications": resolve_publications(calc, sim),
    }
    path.write_text(json.dumps(projection, indent=2) + "\n", encoding="utf-8")
    return projection
