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

from dataclasses import dataclass, field
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

    # Order matters, and mirrors the VBA. Each value is bounded BEFORE any arithmetic
    # that depends on it, so an oversized entered value is rejected with a message
    # rather than overflowing the calculation that would have caught it.
    duration_bounded = True
    if entered.duration < 1:
        problems.append(f"Project Duration must be at least 1 year, not {entered.duration}.")
        duration_bounded = False
    elif entered.duration > limits.max_generated_year_columns:
        # The Architecture Lock structural protection on generated PROJECT-YEAR
        # columns. Independent of the calendar-year window below.
        problems.append(
            f"Project Duration {entered.duration} would generate {entered.duration} "
            f"project-year columns, beyond the structural protection limit of "
            f"{limits.max_generated_year_columns}."
        )
        duration_bounded = False

    years_bounded = True
    for label, value in (
        ("Base Year", entered.base_year),
        ("Project Start Year", entered.start_year),
    ):
        if not limits.min_year <= value <= limits.max_year:
            problems.append(
                f"{label} {value} is outside the supported range "
                f"{limits.min_year}-{limits.max_year}."
            )
            years_bounded = False

    if years_bounded and entered.base_year > entered.start_year:
        problems.append(
            f"Base Year {entered.base_year} is later than Project Start Year "
            f"{entered.start_year}. Costs cannot be priced after the project begins."
        )

    if duration_bounded and years_bounded:
        last = entered.start_year + entered.duration - 1
        if last > limits.max_year:
            problems.append(
                f"Last Project Year would be {last}, beyond the supported structural "
                f"year boundary {limits.max_year}."
            )

    return problems


# ---------------------------------------------------------------------------
# profiling: anchored by project-year index
# ---------------------------------------------------------------------------
def is_data(value: object) -> bool:
    """True when a cell holds something whose loss is real data loss.

    Blank is not data: the user has entered nothing. Numeric zero is not data
    either: 0% is the value a new project-year cell is created with, so removing
    one destroys nothing the user chose.

    EVERYTHING ELSE IS. A non-zero percentage obviously, but also text, an Excel
    error value, or anything else that ended up in the cell by paste. Treating only
    numeric non-zero cells as data would let a shrink silently delete a pasted
    string the user was about to fix, without a destructive warning.
    """
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    if isinstance(value, bool):
        return True
    if isinstance(value, (int, float)):
        return value != 0
    return True


def remap_profiling(
    rows: Mapping[str, Sequence[object]], new_duration: int, initial: float = 0.0
) -> dict[str, list[object]]:
    """Reshape profiling rows to *new_duration* project-year columns.

    Values are preserved by POSITION, because a profiling percentage means 'this
    share falls in project year N', not 'this share falls in 2029'. Growing appends
    new cells at *initial* (0%), so a row that totalled 100% still totals 100% and
    no redistribution or normalisation occurs. Shrinking truncates from the tail.

    A surviving position keeps its value EXACTLY, including blank. Structural
    synchronisation must not quietly turn a blank into 0%: a blank profiling cell is
    invalid data and Model Check has to be able to see it.
    """
    if new_duration < 0:
        raise ValueError(f"new_duration must not be negative, got {new_duration}")
    reshaped: dict[str, list[object]] = {}
    for key, values in rows.items():
        kept = list(values[:new_duration])
        kept.extend([initial] * (new_duration - len(kept)))
        reshaped[key] = kept
    return reshaped


def sync_profiling_values(
    existing: Mapping[str, Sequence[object]],
    keys: Sequence[str],
    year_count: int,
    initial: float = 0.0,
) -> dict[str, list[object]]:
    """Rebuild profiling rows for *keys*, preserving every existing value exactly.

    The distinction this encodes is the one row synchronisation kept getting wrong:

      genuinely NEW driver          -> every project year starts at *initial* (0%)
      genuinely NEW project year    -> that cell starts at *initial* (0%)
      EXISTING (id, year index)     -> preserved exactly, INCLUDING BLANK

    Filling an existing blank with 0% would repair invalid user data inside a
    structural operation and hide it from the Model Check phase whose job is to
    report it.
    """
    rebuilt: dict[str, list[object]] = {}
    for key in keys:
        if key in existing:
            values = list(existing[key][:year_count])
            values.extend([initial] * (year_count - len(values)))
        else:
            values = [initial] * year_count
        rebuilt[key] = values
    return rebuilt


def removed_profiling_values(
    rows: Mapping[str, Sequence[object]], new_duration: int
) -> list[tuple[str, int, object]]:
    """The (id, project-year index, value) triples a shrink would destroy.

    Reported when ``is_data`` holds: blank and numeric zero destroy nothing, and
    anything else does. Project-year indices are 1-based, matching the user's view.
    """
    destroyed: list[tuple[str, int, object]] = []
    for key, values in rows.items():
        for offset, value in enumerate(values):
            if offset >= new_duration and is_data(value):
                destroyed.append((key, offset + 1, value))
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
    rows: Mapping[str, Mapping[int, float | None]],
    new_years: Iterable[int],
    surviving_profiles: Iterable[str] | None = None,
) -> list[tuple[str, int, object]]:
    """The (profile, calendar year, rate) triples a synchronisation would destroy.

    TWO INDEPENDENT LOSS MECHANISMS, and a rate is lost if EITHER applies:

      * its calendar year leaves the required span, or
      * its profile name leaves the Config master list.

    The second is not a timeline change at all: deleting a profile from Config
    destroys that row's rates on the next synchronisation even when Base Year,
    Start Year and Duration are untouched.

    Each (profile, year) cell is judged ONCE, so a cell whose profile and whose year
    both disappear in the same operation is counted once, not twice.

    ``surviving_profiles`` of None means "profiles are not being changed", which is
    different from an empty collection meaning "every profile has been removed".
    Only non-blank rates count: a blank leaving destroys nothing.
    """
    surviving_years = set(new_years)
    keep_profiles = None if surviving_profiles is None else set(surviving_profiles)

    destroyed: list[tuple[str, int, object]] = []
    for key, rates in rows.items():
        profile_lost = keep_profiles is not None and key not in keep_profiles
        for year, rate in rates.items():
            if not is_data(rate):
                continue
            if profile_lost or year not in surviving_years:
                destroyed.append((key, year, rate))
    return destroyed


def removed_profiles(
    rows: Mapping[str, Mapping[int, float | None]], surviving_profiles: Iterable[str]
) -> list[str]:
    """Inflation rows whose profile name has left the Config master list."""
    keep = set(surviving_profiles)
    return [key for key in rows if key not in keep]


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
    profiling_losses: list[tuple[str, str, int, object]]
    inflation_losses: list[tuple[str, int, object]]
    removed_profile_names: list[str] = field(default_factory=list)

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
    profiling: Mapping[str, Mapping[str, Sequence[object]]],
    inflation: Mapping[str, Mapping[int, float | None]],
    surviving_profiles: Iterable[str] | None = None,
) -> DestructiveImpact:
    """What a synchronisation would destroy, given the current grid contents.

    ``profiling`` is {grid name: {permanent id: [values by project-year index]}}.
    ``surviving_profiles`` is the current Config profile master; pass it whenever
    profile rows will be synchronised, which Apply always does. Omitting it means
    "profiles are not changing" and suppresses that half of the assessment.

    Runs before any modification, which is why cancelling needs no rollback: the
    user is asked before a single cell has moved.
    """
    new_duration = change.new.duration or 0
    profiling_losses: list[tuple[str, str, int, object]] = []
    for grid_name, rows in profiling.items():
        for key, index, value in removed_profiling_values(rows, new_duration):
            profiling_losses.append((grid_name, key, index, value))

    return DestructiveImpact(
        removed_project_year_indices=change.removed_project_year_indices,
        removed_inflation_years=change.removed_inflation_years,
        profiling_losses=profiling_losses,
        inflation_losses=removed_inflation_values(
            inflation, change.new_inflation_years, surviving_profiles
        ),
        removed_profile_names=(
            [] if surviving_profiles is None else removed_profiles(inflation, surviving_profiles)
        ),
    )


def apply_change(
    change: TimelineChange,
    profiling: Mapping[str, Mapping[str, Sequence[object]]],
    inflation: Mapping[str, Mapping[int, float | None]],
    initial_profile_value: float = 0.0,
    surviving_profiles: Iterable[str] | None = None,
) -> tuple[dict[str, dict[str, list[object]]], dict[str, dict[int, float | None]]]:
    """The grid contents after *change*, as one combined transition.

    When ``surviving_profiles`` is given, rows whose profile has left the Config
    master are dropped, which is the change the destructive assessment warned about.
    """
    new_duration = change.new.duration or 0
    reshaped_profiling = {
        grid_name: remap_profiling(rows, new_duration, initial_profile_value)
        for grid_name, rows in profiling.items()
    }
    kept = inflation
    if surviving_profiles is not None:
        keep = set(surviving_profiles)
        kept = {name: rates for name, rates in inflation.items() if name in keep}
    return reshaped_profiling, remap_inflation(kept, change.new_inflation_years)
