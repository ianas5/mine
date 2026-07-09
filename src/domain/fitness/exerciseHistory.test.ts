import { summarizeExerciseHistory, type HistorySetRow } from './exerciseHistory';

const row = (
  workoutId: string,
  order: number,
  weightKg: number,
  reps: number,
  warmup = false,
): HistorySetRow => ({
  workoutId,
  date: '2026-07-08',
  workoutOrder: order,
  weightKg,
  reps,
  warmup,
});

describe('summarizeExerciseHistory (UI_UX §4.1)', () => {
  it('returns an all-null preview for a first-ever exercise (P8)', () => {
    expect(summarizeExerciseHistory([], 'external')).toEqual({
      last: null,
      bestWeightSet: null,
      bestE1rmKg: null,
    });
  });

  it('takes Last from the most recent workout and excludes warm-ups', () => {
    const rows = [
      row('w2', 200, 80, 8),
      row('w2', 200, 80, 7),
      row('w2', 200, 60, 10, true), // warm-up excluded
      row('w1', 100, 70, 10),
    ];
    const preview = summarizeExerciseHistory(rows, 'external');

    expect(preview.last?.sets).toEqual([
      { weightKg: 80, reps: 8 },
      { weightKg: 80, reps: 7 },
    ]);
  });

  it('computes best weight set and best e1RM across all history', () => {
    const rows = [row('w2', 200, 80, 8), row('w1', 100, 85, 6), row('w1', 100, 100, 15)];
    const preview = summarizeExerciseHistory(rows, 'external');

    expect(preview.bestWeightSet).toEqual({ weightKg: 100, reps: 15 });
    // 100kg×15 is not e1RM-eligible (>12 reps); best eligible is 85×6.
    expect(preview.bestE1rmKg).toBeCloseTo(85 * (1 + 6 / 30), 4);
  });
});
