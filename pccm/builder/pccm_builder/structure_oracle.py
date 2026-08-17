"""Pure structural semantics for the Phase-4 timeline.

This module is the *specification* of what Apply / Update Timeline does to the
structural grids. It is pure Python over plain data: no Excel, no COM, no VBA and
no I/O, so every rule below can be tested exhaustively on Linux.

It does NOT prove the VBA is correct. It defines the transformation the VBA must
implement, and it generates the expected values the Windows functional harness
compares the real workbook against, so the two cannot drift apart silently. Only a
clean Windows run proves the runtime.

The two anchoring invariants are deliberately different, and everything here turns
on the distinction:

  profiling  anchored by PROJECT-YEAR INDEX. Moving the start year relabels the
             headers and moves nothing. Only a duration change adds or removes
             cells, and only at the tail.
  inflation  anchored by CALENDAR YEAR. A rate entered against 2029 stays attached
             to 2029 for as long as 2029 is inside the required span, whatever
             column index it lands in.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Mapping, Sequence

BLANK = None


# ---------------------------------------------------------------------------
# timeline
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Timeline:
    """A timeline triple. Any member may be None, meaning 'not entered'."""

    base_year: int | None
    start_year: int | None
    duration: int | None

    @property
    def is_blank(self) -> bool:
        return self.base_year is None and self.start_year is None and self.duration is None

    @property
    def is_complete(self) -> bool:
        return None not in (self.base_year, self.start_year, self.duration)

    @property
    def last_project_year(self) -> int | None:
        if self.start_year is None or self.duration is None:
            return None
        return self.start_year + self.duration - 1

    @property
    def project_years(self) -> list[int]:
        """Calendar year for project-year index 1..duration."""
        if not self.is_complete:
            return []
        return [self.start_year + offset for offset in range(self.duration)]

    @property
    def inflation_years(self) -> list[int]:
        """Required escalation years: applied base year + 1 through last project year.

        Legitimately empty when base year == start year and duration == 1: there is
        no year between the price base and the end of the project, so no escalation
        assumption is required. That is a valid model, not a malformed one.
        """
        if not self.is_complete:
            return []
        first = self.base_year + 1
        last = self.last_project_year
        if first > last:
            return []
        return list(range(first, last + 1))

    @property
    def has_empty_inflation_span(self) -> bool:
        return self.is_complete and not self.inflation_years


@dataclass(frozen=True)
class Limits:
    min_year: int
    max_year: int
    max_generated_year_columns: int


def prevalidate(entered: Timeline, limits: Limits) -> list[str]:
    """Every reason *entered* may not be applied. Empty means it may.

    Returns all failures rather than the first, so a user fixing one is not sent
    back for the next. No partial application is ever permitted, so this runs to
    completion before a single structural cell is touched.
    """
    problems: list[str] = []

    for label, value in (
        ("Base Year", entered.base_year),
        ("Project Start Year", entered.start_year),
        ("Project Duration (Years)", entered.duration),
    ):
        if value is None:
            problems.append(f"{label} is blank; all three timeline values are required.")
        elif isinstance(value, bool) or not isinstance(value, int):
            problems.append(f"{label} must be a whole number, not {value!r}.")

    if problems:
        return problems

    if entered.duration < 1:
        problems.append(f"Project Duration must be at least 1 year, not {entered.duration}.")
    for label, value in (("Base Year", entered.base_year), ("Project Start Year", entered.start_year)):
        if not limits.min_year <= value <= limits.max_year:
            problems.append(
                f"{label} {value} is outside the supported range "
                f"{limits.min_year}-{limits.max_year}."
            )
    if entered.base_year > entered.start_year:
        problems.append(
            f"Base Year {entered.base_year} is later than Project Start Year "
            f"{entered.start_year}. Costs cannot be priced after the project begins."
        )

    if entered.duration is not None and entered.duration >= 1 and entered.start_year is not None:
        last = entered.start_year + entered.duration - 1
        if last > limits.max_year:
            problems.append(
                f"Last Project Year would be {last}, beyond the supported structural "
                f"year boundary {limits.max_year}."
            )
        if entered.duration > limits.max_generated_year_columns:
            problems.append(
                f"Project Duration {entered.duration} would generate "
                f"{entered.duration} project-year columns, beyond the structural "
                f"protection limit of {limits.max_generated_year_columns}."
            )

    return problems


# ---------------------------------------------------------------------------
# profiling: anchored by project-year index
# ---------------------------------------------------------------------------
def remap_profiling(
    rows: Mapping[str, Sequence[float]], new_duration: int, initial: float = 0.0
) -> dict[str, list[float]]:
    """Reshape profiling rows to *new_duration* project-year columns.

    Values are preserved by POSITION, because a profiling percentage means 'this
    share falls in project year N', not 'this share falls in 2029'. Growing appends
    new cells at *initial* (0%), so a row that totalled 100% still totals 100% and
    no redistribution or normalisation occurs. Shrinking truncates from the tail.
    """
    if new_duration < 0:
        raise ValueError(f"new_duration must not be negative, got {new_duration}")
    reshaped: dict[str, list[float]] = {}
    for key, values in rows.items():
        kept = list(values[:new_duration])
        kept.extend([initial] * (new_duration - len(kept)))
        reshaped[key] = kept
    return reshaped


def removed_profiling_values(
    rows: Mapping[str, Sequence[float]], new_duration: int
) -> list[tuple[str, int, float]]:
    """The (id, project-year index, value) triples a shrink would destroy.

    Only non-zero values are reported: removing a 0% cell destroys no user data and
    needs no destructive warning beyond the ordinary structural confirmation.
    Project-year indices are 1-based, matching what the user sees.
    """
    destroyed: list[tuple[str, int, float]] = []
    for key, values in rows.items():
        for offset, value in enumerate(values):
            if offset >= new_duration and value:
                destroyed.append((key, offset + 1, float(value)))
    return destroyed


# ---------------------------------------------------------------------------
# inflation: anchored by calendar year
# ---------------------------------------------------------------------------
def remap_inflation(
    rows: Mapping[str, Mapping[int, float | None]], new_years: Sequence[int]
) -> dict[str, dict[int, float | None]]:
    """Reshape inflation rows to exactly *new_years*, keyed by calendar year.

    A rate survives if and only if its calendar year survives, whatever column it
    used to sit in. Newly required years arrive BLANK: a rate the user has not
    entered must never be fabricated as zero, because a missing rate is a blocking
    Model Check condition later and a fabricated zero would hide it forever.
    """
    reshaped: dict[str, dict[int, float | None]] = {}
    for key, rates in rows.items():
        reshaped[key] = {year: rates.get(year, BLANK) for year in new_years}
    return reshaped


def removed_inflation_values(
    rows: Mapping[str, Mapping[int, float | None]], new_years: Iterable[int]
) -> list[tuple[str, int, float]]:
    """The (profile, calendar year, rate) triples a remap would destroy.

    Only non-blank rates are reported. A blank leaving the span destroys nothing.
    """
    surviving = set(new_years)
    destroyed: list[tuple[str, int, float]] = []
    for key, rates in rows.items():
        for year, rate in rates.items():
            if year not in surviving and rate is not None:
                destroyed.append((key, year, float(rate)))
    return destroyed


# ---------------------------------------------------------------------------
# row synchronisation: keyed by permanent identity, never by row position
# ---------------------------------------------------------------------------
def sync_rows(
    keys: Sequence[str], existing: Mapping[str, object]
) -> tuple[list[str], list[str], list[str]]:
    """(ordered keys, added, removed) for a grid synchronised against *keys*.

    Ownership follows the permanent key: reordering the driver table reorders the
    grid but transfers nothing between rows, and a key that survives keeps its data
    whatever row it now occupies.
    """
    wanted = [k for k in keys if k]
    duplicates = sorted({k for k in wanted if wanted.count(k) > 1})
    if duplicates:
        raise ValueError(f"duplicate permanent identifiers: {duplicates}")
    added = [k for k in wanted if k not in existing]
    removed = [k for k in existing if k not in set(wanted)]
    return wanted, added, removed


# ---------------------------------------------------------------------------
# permanent identity
# ---------------------------------------------------------------------------
def allocate_id(counter: int, prefix: str, pad_width: int) -> tuple[str, int]:
    """Allocate the next permanent identifier. Returns (id, advanced counter).

    The counter is the only source of a sequence number. It is never derived from a
    row number or a row count, and deletion never decrements it, so an identifier is
    never reused. Pad width is a minimum display width: once the sequence passes it,
    identifiers simply get longer (CL-999 -> CL-1000) rather than hitting a ceiling.
    """
    if counter < 0:
        raise ValueError(f"counter must not be negative, got {counter}")
    nxt = counter + 1
    return f"{prefix}{nxt:0{pad_width}d}", nxt


def highest_sequence(ids: Iterable[str], prefix: str) -> int:
    """The largest sequence number present in *ids*. Zero when there are none.

    Used by structural revalidation: a counter that has fallen below an identifier
    it already issued would reissue that identifier next time, so the check is
    'counter >= highest issued', never 'counter == count of rows'.
    """
    highest = 0
    for value in ids:
        if not value or not str(value).startswith(prefix):
            continue
        tail = str(value)[len(prefix):]
        if tail.isdigit():
            highest = max(highest, int(tail))
    return highest


# ---------------------------------------------------------------------------
# one combined transition
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class TimelineChange:
    """One coherent old -> new transition.

    Base year, start year and duration may all change before Apply. The delta is
    computed once, from one old configuration to one new configuration, and applied
    as a single transition. Three sequential partial transitions would expose
    half-applied structural states and could raise a destructive warning for a cell
    the combined change never actually removes.
    """

    old: Timeline
    new: Timeline

    @property
    def old_project_years(self) -> list[int]:
        return self.old.project_years

    @property
    def new_project_years(self) -> list[int]:
        return self.new.project_years

    @property
    def old_inflation_years(self) -> list[int]:
        return self.old.inflation_years

    @property
    def new_inflation_years(self) -> list[int]:
        return self.new.inflation_years

    @property
    def duration_delta(self) -> int:
        return (self.new.duration or 0) - (self.old.duration or 0)

    @property
    def removed_project_year_indices(self) -> list[int]:
        """1-based project-year indices that disappear. Empty unless duration shrinks."""
        old_count = self.old.duration or 0
        new_count = self.new.duration or 0
        return list(range(new_count + 1, old_count + 1))

    @property
    def removed_inflation_years(self) -> list[int]:
        surviving = set(self.new_inflation_years)
        return [y for y in self.old_inflation_years if y not in surviving]

    @property
    def added_inflation_years(self) -> list[int]:
        previous = set(self.old_inflation_years)
        return [y for y in self.new_inflation_years if y not in previous]

    @property
    def headers_relabelled(self) -> bool:
        """True when profiling headers change but no profiling cell moves.

        The pure start-year shift: same duration, different calendar labels.
        """
        return (
            self.old.is_complete
            and self.new.is_complete
            and self.old.duration == self.new.duration
            and self.old.start_year != self.new.start_year
        )


@dataclass(frozen=True)
class DestructiveImpact:
    """What a transition would permanently delete, computed BEFORE anything moves."""

    removed_project_year_indices: list[int]
    removed_inflation_years: list[int]
    profiling_losses: list[tuple[str, str, int, float]]
    inflation_losses: list[tuple[str, int, float]]

    @property
    def profiling_loss_count(self) -> int:
        return len(self.profiling_losses)

    @property
    def inflation_loss_count(self) -> int:
        return len(self.inflation_losses)

    @property
    def is_destructive(self) -> bool:
        """True only when real user data would be lost.

        Removing 0% profiling cells or blank inflation cells destroys nothing, so it
        needs no destructive warning beyond the ordinary structural confirmation.
        """
        return bool(self.profiling_losses or self.inflation_losses)

    def affected_ids(self, limit: int = 5) -> list[str]:
        """Representative permanent IDs, for the confirmation prompt."""
        seen: list[str] = []
        for _grid, key, _index, _value in self.profiling_losses:
            if key not in seen:
                seen.append(key)
            if len(seen) >= limit:
                break
        return seen


def assess(
    change: TimelineChange,
    profiling: Mapping[str, Mapping[str, Sequence[float]]],
    inflation: Mapping[str, Mapping[int, float | None]],
) -> DestructiveImpact:
    """What *change* would destroy, given the current grid contents.

    ``profiling`` is {grid name: {permanent id: [values by project-year index]}}.
    Runs before any modification, which is why cancelling needs no rollback: the
    user is asked before a single cell has moved.
    """
    new_duration = change.new.duration or 0
    profiling_losses: list[tuple[str, str, int, float]] = []
    for grid_name, rows in profiling.items():
        for key, index, value in removed_profiling_values(rows, new_duration):
            profiling_losses.append((grid_name, key, index, value))

    return DestructiveImpact(
        removed_project_year_indices=change.removed_project_year_indices,
        removed_inflation_years=change.removed_inflation_years,
        profiling_losses=profiling_losses,
        inflation_losses=removed_inflation_values(inflation, change.new_inflation_years),
    )


def apply_change(
    change: TimelineChange,
    profiling: Mapping[str, Mapping[str, Sequence[float]]],
    inflation: Mapping[str, Mapping[int, float | None]],
    initial_profile_value: float = 0.0,
) -> tuple[dict[str, dict[str, list[float]]], dict[str, dict[int, float | None]]]:
    """The grid contents after *change*, as one combined transition."""
    new_duration = change.new.duration or 0
    reshaped_profiling = {
        grid_name: remap_profiling(rows, new_duration, initial_profile_value)
        for grid_name, rows in profiling.items()
    }
    return reshaped_profiling, remap_inflation(inflation, change.new_inflation_years)
