import type { ExerciseSetRow } from '@/domain/fitness';

import { computeExerciseReport, EMPTY_REPORT } from './exerciseReport';

const row = (
  workoutId: string,
  weightKg: number,
  reps: number,
  order: number,
  warmup = false,
): ExerciseSetRow => ({
  workoutId,
  date: order === 2 ? '2026-07-05' : '2026-07-01',
  workoutOrder: order,
  weightKg,
  reps,
  warmup,
  counting: 'none',
});

describe('computeExerciseReport (ANALYTICS §5.5)', () => {
  it('returns the empty report when nothing has been logged', () => {
    expect(computeExerciseReport([], 'external', null)).toEqual(EMPTY_REPORT);
  });

  it('aggregates sessions, sets, volume, averages, bests and last performed', () => {
    const rows = [
      row('w1', 100, 10, 1),
      row('w1', 100, 8, 1),
      row('w1', 60, 12, 1, true), // warm-up: excluded from every metric
      row('w2', 120, 5, 2),
    ];
    const report = computeExerciseReport(rows, 'external', null);

    expect(report.totalSessions).toBe(2);
    expect(report.totalWorkingSets).toBe(3);
    expect(report.totalVolumeKg).toBe(100 * 10 + 100 * 8 + 120 * 5); // 2400
    expect(report.avgRepsPerWorkingSet).toBeCloseTo((10 + 8 + 5) / 3, 4);
    expect(report.avgEffectiveLoadKg).toBeCloseTo((100 + 100 + 120) / 3, 4);
    expect(report.bests.heaviestWeightKg).toBe(120);
    expect(report.lastPerformed?.date).toBe('2026-07-05'); // the most recent workout
  });
});
