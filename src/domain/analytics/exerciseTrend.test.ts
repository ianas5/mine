import type { ExerciseSetRow } from '@/domain/fitness';

import { bestE1rmSeries, computeExerciseTrend } from './exerciseReport';

/** One working set in a workout on `date`; distinct workoutIds become distinct sessions. */
const row = (workoutId: string, date: string, weightKg: number, reps: number): ExerciseSetRow => ({
  workoutId,
  date,
  workoutOrder: Number(workoutId.slice(1)),
  weightKg,
  reps,
  warmup: false,
  counting: 'none',
});

describe('bestE1rmSeries (ANALYTICS §5.1 strength series)', () => {
  it('emits one point per workout = its best e1RM, ascending', () => {
    const rows = [
      row('w1', '2026-01-01', 100, 5),
      row('w1', '2026-01-01', 90, 5), // same workout, lower e1RM → not the point
      row('w2', '2026-01-15', 105, 5),
    ];
    const series = bestE1rmSeries(rows, 'external', null);
    expect(series).toHaveLength(2);
    expect(series[0]!.date).toBe('2026-01-01');
    expect(series[1]!.value).toBeGreaterThan(series[0]!.value);
  });

  it('ignores e1RM-ineligible high-rep sets', () => {
    // 20 reps is above the e1RM eligibility cap → no eligible set → no point.
    expect(bestE1rmSeries([row('w1', '2026-01-01', 60, 20)], 'external', null)).toEqual([]);
  });
});

describe('computeExerciseTrend', () => {
  it('reports an increasing, improving strength trend over enough sessions', () => {
    const rows = [
      row('w1', '2026-01-01', 100, 5),
      row('w2', '2026-01-15', 105, 5),
      row('w3', '2026-01-29', 110, 5),
    ];
    const { series, trend } = computeExerciseTrend(rows, 'external', null, '2026-02-01');
    expect(series).toHaveLength(3);
    expect(trend.status).toBe('ok');
    if (trend.status !== 'ok') return;
    expect(trend.value.direction).toBe('increasing');
    expect(trend.value.classification).toBe('improving'); // higher e1RM is good
    expect(trend.value.slopePerWeek).toBeGreaterThan(0); // progression rate
  });

  it('propagates insufficient-data below the §6.4 minimums', () => {
    const { trend } = computeExerciseTrend(
      [row('w1', '2026-01-01', 100, 5), row('w2', '2026-01-15', 105, 5)],
      'external',
      null,
      '2026-02-01',
    );
    expect(trend.status).toBe('insufficient-data');
  });
});
