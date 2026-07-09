import {
  countableWorkoutsByWeek,
  currentWeekProgress,
  weekConsistencyPercent,
  weekStartIso,
  weeklyStreak,
} from './consistency';

// 2026-03-09 is a Monday; 2026-03-15 the Sunday of the same ISO week.
describe('weekStartIso (ISO Monday, §3.8)', () => {
  it('maps every day of a week back to its Monday', () => {
    expect(weekStartIso('2026-03-09')).toBe('2026-03-09'); // Monday → itself
    expect(weekStartIso('2026-03-15')).toBe('2026-03-09'); // Sunday → same Monday
    expect(weekStartIso('2026-03-16')).toBe('2026-03-16'); // next Monday
  });
});

describe('countableWorkoutsByWeek', () => {
  it('counts each countable workout in its ISO week (two on one day = 2)', () => {
    const byWeek = countableWorkoutsByWeek([
      '2026-03-09',
      '2026-03-09',
      '2026-03-15',
      '2026-03-16',
    ]);
    expect(byWeek.get('2026-03-09')).toBe(3); // Mon + Mon + Sun of that week
    expect(byWeek.get('2026-03-16')).toBe(1);
  });
});

describe('currentWeekProgress (partial week = progress, §4)', () => {
  it('reports completed of planned for the current week only', () => {
    const dates = ['2026-03-09', '2026-03-11', '2026-03-02']; // last is prior week
    expect(currentWeekProgress(dates, '2026-03-12', 4)).toEqual({ completed: 2, planned: 4 });
  });

  it('a fresh Monday reads 0 of N, not a crash', () => {
    expect(currentWeekProgress(['2026-03-02'], '2026-03-09', 4)).toEqual({
      completed: 0,
      planned: 4,
    });
  });
});

describe('weekConsistencyPercent (completed weeks, §3.8)', () => {
  it('is completed / planned, capped at 100', () => {
    expect(weekConsistencyPercent(2, 4)).toBe(50);
    expect(weekConsistencyPercent(5, 4)).toBe(100); // over-delivery caps at 100
    expect(weekConsistencyPercent(1, 0)).toBe(0); // no plan → 0, never divide-by-zero
  });
});

describe('weeklyStreak (§3.8 weekly, not daily)', () => {
  const target = 3;

  it('counts consecutive completed weeks that meet target', () => {
    // Weeks of Mar 9, Mar 2, Feb 23 each have ≥ 3; today is mid-week Mar 12 with 3 already.
    const dates = [
      '2026-03-09',
      '2026-03-10',
      '2026-03-11', // current week: 3
      '2026-03-02',
      '2026-03-03',
      '2026-03-04', // prior week: 3
      '2026-02-23',
      '2026-02-24',
      '2026-02-25', // week before: 3
    ];
    expect(weeklyStreak(dates, '2026-03-12', target)).toBe(3);
  });

  it('does not break the streak when the current week is only partway there', () => {
    // Current week has just 1 (in progress); prior two completed weeks met target.
    const dates = [
      '2026-03-09', // current week: 1 (partial, not a failure)
      '2026-03-02',
      '2026-03-03',
      '2026-03-04', // prior: 3
      '2026-02-23',
      '2026-02-24',
      '2026-02-25', // before: 3
    ];
    expect(weeklyStreak(dates, '2026-03-10', target)).toBe(2);
  });

  it('breaks on a fully-elapsed week that missed target', () => {
    const dates = [
      '2026-03-09',
      '2026-03-10',
      '2026-03-11', // current: 3 (counts)
      '2026-03-02',
      '2026-03-03', // prior: only 2 → miss, streak stops here
      '2026-02-23',
      '2026-02-24',
      '2026-02-25', // earlier meets, but unreachable past the miss
    ];
    expect(weeklyStreak(dates, '2026-03-12', target)).toBe(1);
  });

  it('is zero with no history or a non-positive target', () => {
    expect(weeklyStreak([], '2026-03-12', target)).toBe(0);
    expect(weeklyStreak(['2026-03-09'], '2026-03-12', 0)).toBe(0);
  });
});
