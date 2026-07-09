import { addDaysIso, isoWeekday, type IsoDate } from '@/core/utils';

/**
 * Consistency & streaks (FITNESS_DOMAIN §3.8). Weekly, not daily — daily streaks punish
 * normal rest days. All pure over the dates of **countable** workouts (a workout with
 * ≥ 1 working set); the caller filters countable via `isCountableWorkout`. Phase 17's
 * `WorkoutAnalyticsCalculator` consumes these same functions.
 */

/** The ISO-Monday week start for a date (0 = Monday … 6 = Sunday). */
export function weekStartIso(date: IsoDate): IsoDate {
  return addDaysIso(date, -isoWeekday(date));
}

/** Countable workouts grouped by ISO week (keyed by the week's Monday). */
export function countableWorkoutsByWeek(dates: readonly IsoDate[]): Map<IsoDate, number> {
  const byWeek = new Map<IsoDate, number>();
  for (const date of dates) {
    const week = weekStartIso(date);
    byWeek.set(week, (byWeek.get(week) ?? 0) + 1);
  }
  return byWeek;
}

export interface WeekProgress {
  readonly completed: number;
  readonly planned: number;
}

/**
 * The in-progress week as **progress**, never a percentage (§4 partial-week rule): a
 * fresh Monday reads "0 of N", not a consistency crash. `planned` is the program's
 * sessions for the week, else `weeklyWorkoutTarget` (resolved by the caller).
 */
export function currentWeekProgress(
  dates: readonly IsoDate[],
  today: IsoDate,
  planned: number,
): WeekProgress {
  const week = weekStartIso(today);
  return { completed: countableWorkoutsByWeek(dates).get(week) ?? 0, planned };
}

/** Completed-week consistency % (§3.8) — only meaningful for a fully-elapsed week. */
export function weekConsistencyPercent(completed: number, planned: number): number {
  if (planned <= 0) return 0;
  return Math.min(100, Math.round((completed / planned) * 100));
}

/**
 * Training streak (§3.8): consecutive weeks meeting ≥ `target` countable workouts,
 * counting back from today. The current (partial) week counts **only if it already
 * meets target**, and never breaks the streak; any fully-elapsed prior week below
 * target ends it. Zero-workout weeks (count 0) end the streak.
 */
export function weeklyStreak(dates: readonly IsoDate[], today: IsoDate, target: number): number {
  if (target <= 0) return 0;
  const byWeek = countableWorkoutsByWeek(dates);

  let streak = 0;
  let week = weekStartIso(today);
  // Current week: contributes only when already met; otherwise it's still in progress.
  if ((byWeek.get(week) ?? 0) >= target) streak += 1;
  week = addDaysIso(week, -7);

  while ((byWeek.get(week) ?? 0) >= target) {
    streak += 1;
    week = addDaysIso(week, -7);
  }
  return streak;
}
