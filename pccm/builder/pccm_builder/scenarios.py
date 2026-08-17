"""Phase-4 Windows functional scenarios, derived from the structural oracle.

Emitted to ``build/phase4_scenarios.json`` and read at runtime by the Windows
functional harness, so the harness contains no hand-written expected values.

Why this exists: the shape a timeline produces -- which project-year headers,
which calendar years, how many columns, whether the inflation span is empty -- is
already defined once, in ``structure_oracle``. Restating those numbers inside a
PowerShell script would create a second definition that could silently disagree
with the first. Here the oracle computes them, a Linux test asserts the emitted
fixture still matches the oracle, and Windows compares the real workbook against
the same numbers.

The harness still does its own before/after comparisons for value preservation:
"unchanged bit for bit" is a statement about the workbook's own previous state,
not about a constant, so hard-coding expected percentages would test the wrong
thing.
"""

from __future__ import annotations

from .structure_oracle import Limits, Timeline, TimelineChange


def _shape(timeline: Timeline) -> dict:
    return {
        "applied": {
            "base_year": timeline.base_year,
            "start_year": timeline.start_year,
            "duration": timeline.duration,
        },
        "last_project_year": timeline.last_project_year,
        "profiling_headers": [str(year) for year in timeline.project_years],
        "profiling_year_count": len(timeline.project_years),
        "inflation_years": [str(year) for year in timeline.inflation_years],
        "inflation_year_count": len(timeline.inflation_years),
        "inflation_span_is_empty": timeline.has_empty_inflation_span,
    }


def _transition(previous: Timeline, current: Timeline) -> dict:
    change = TimelineChange(old=previous, new=current)
    return {
        "removed_project_year_indices": change.removed_project_year_indices,
        "removed_inflation_years": [str(y) for y in change.removed_inflation_years],
        "added_inflation_years": [str(y) for y in change.added_inflation_years],
        "duration_delta": change.duration_delta,
        "headers_relabelled_only": change.headers_relabelled,
    }


def build_scenarios(limits: Limits) -> dict:
    """The ordered Phase-4 functional scenarios, D through J of the test matrix.

    Each step is applied to the workbook in sequence, so ``previous`` is the state
    the step starts from. ``confirm`` is the reply the harness gives to the
    confirmation prompt; the shrink step is deliberately run twice, once cancelled
    and once accepted, because the cancellation path is the one that proves nothing
    moved before the user was asked.
    """
    steps: list[dict] = []

    def add(
        key: str,
        title: str,
        previous: Timeline,
        entered: Timeline,
        *,
        confirm: bool = True,
        expect_destructive: bool = False,
        expect_rejected: bool = False,
        note: str = "",
    ) -> None:
        step = {
            "key": key,
            "title": title,
            "entered": {
                "base_year": entered.base_year,
                "start_year": entered.start_year,
                "duration": entered.duration,
            },
            "confirm": confirm,
            "expect_destructive_prompt": expect_destructive,
            "expect_rejected": expect_rejected,
            "note": note,
        }
        # A rejected or cancelled step must leave the PREVIOUS shape in place; that
        # is the whole assertion, so the expected shape is the previous one.
        step["expect"] = _shape(previous if (expect_rejected or not confirm) else entered)
        step["transition"] = _transition(previous, entered)
        steps.append(step)

    blank = Timeline(None, None, None)

    first = Timeline(base_year=2026, start_year=2028, duration=3)
    add("D_first_apply", "First timeline application", blank, first,
        note="No timeline is applied yet, so nothing can be destroyed.")

    longer = Timeline(base_year=2026, start_year=2028, duration=5)
    add("E_duration_increase", "Duration increase 3 -> 5", first, longer,
        note="Existing percentages must be unchanged bit for bit; the two new "
             "project-year cells arrive at 0%; new inflation years arrive blank.")

    shifted = Timeline(base_year=2026, start_year=2030, duration=5)
    add("F_start_year_shift", "Start-year shift 2028 -> 2030, same duration", longer, shifted,
        note="Profiling values stay in the same project-year positions and only the "
             "headers relabel; inflation rates survive by calendar-year intersection.")

    shrink = Timeline(base_year=2026, start_year=2030, duration=3)
    add("G_duration_decrease_cancel", "Duration decrease 5 -> 3, CANCELLED", shifted, shrink,
        confirm=False, expect_destructive=True,
        note="Non-zero percentages sit in the years that would be removed, so a "
             "destructive confirmation is required. Cancelling must leave every "
             "structural value logically unchanged, and needs no rollback because "
             "the prompt precedes any modification.")

    add("G_duration_decrease_accept", "Duration decrease 5 -> 3, ACCEPTED", shifted, shrink,
        confirm=True, expect_destructive=True,
        note="The same transition, accepted. The removed project years disappear and "
             "the surviving positions are untouched.")

    base_earlier = Timeline(base_year=2024, start_year=2030, duration=3)
    add("H_base_year_earlier", "Base Year moved earlier 2026 -> 2024", shrink, base_earlier,
        note="Profiling percentages must not change at all. The inflation span grows "
             "at the early end and the new years arrive blank.")

    base_later = Timeline(base_year=2029, start_year=2030, duration=3)
    add("H_base_year_later", "Base Year moved later 2024 -> 2029", base_earlier, base_later,
        expect_destructive=True,
        note="Profiling percentages must not change. Inflation years leave the span at "
             "the early end; any non-blank rate among them makes this destructive.")

    combined = Timeline(base_year=2031, start_year=2033, duration=4)
    add("I_combined_change", "Combined change: base, start and duration together",
        base_later, combined, expect_destructive=True,
        note="One coherent old -> new transition. The workbook must never expose an "
             "intermediate state in which only some of the three had been applied.")

    degenerate = Timeline(base_year=2035, start_year=2035, duration=1)
    add("J_degenerate_span", "Degenerate inflation span: base = start, duration = 1",
        combined, degenerate, expect_destructive=True,
        note="A legitimate model with NO required inflation years. The inflation grid "
             "must keep its fixed column and show the empty-span message; it must not "
             "become a zero-width or malformed table.")

    rejected = Timeline(base_year=2040, start_year=2035, duration=1)
    add("J_rejected_base_after_start", "Rejected: Base Year later than Start Year",
        degenerate, rejected, expect_rejected=True,
        note="Prevalidation must reject this before any structural modification, "
             "leaving the applied triple exactly as it was.")

    return {
        "generated_by": "pccm structure oracle",
        "limits": {
            "min_year": limits.min_year,
            "max_year": limits.max_year,
            "max_generated_year_columns": limits.max_generated_year_columns,
        },
        "identity": {
            "cost_sequence": ["CL-001", "CL-002", "CL-003", "CL-004"],
            "risk_sequence": ["R-001", "R-002", "R-003"],
            "note": "Add three, delete the second, add one more: the fourth add must "
                    "issue CL-004. CL-002 must never be reused.",
        },
        "steps": steps,
    }
