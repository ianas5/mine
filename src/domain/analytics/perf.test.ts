/**
 * Performance watchdog (ANALYTICS §7, IMPLEMENTATION_ROADMAP Phase 22 gate 2). Generates a
 * synthetic **5-year** dataset and times the full dashboard-analytics path (every calculator
 * + the whole insight rule set) and a single chart series. The §7 budgets — dashboard ≤ 50 ms,
 * chart ≤ 16 ms — are stated "on a modern iPhone"; this Node run is a coarse regression guard
 * (asserted against a generous ceiling) that also prints the real numbers for the CP-E report.
 * A device timing-log assertion remains the authoritative on-device check.
 */
import type { IsoDate } from '@/core/utils';
import { addDaysIso } from '@/core/utils';
import type { BodyField, BodySnapshot } from '@/domain/body';
import type { MuscleGroup } from '@/domain/fitness';

import { computeBodyAnalytics } from './bodyAnalytics';
import { bucketSeries } from './bucketing';
import { computeMuscleReports } from './muscleAnalytics';
import { computeNutritionAnalytics, type DailyNutrition } from './nutritionAnalytics';
import { computeRecompSignal } from './recomp';
import { rangeWindow } from './ranges';
import type { SeriesPoint } from './timeSeries';
import type { TrainingWorkout, WeighIn } from './trainingData';
import { computeWorkoutAnalytics } from './workoutAnalytics';

const TODAY: IsoDate = '2026-03-15';
const YEARS = 5;
const DAYS = YEARS * 365;

const GROUPS: MuscleGroup[] = ['chest', 'back', 'quads', 'shoulders', 'biceps', 'triceps'];
const SITE_FIELDS: readonly BodyField[] = ['waistCm', 'chestCm', 'leftArmCm', 'rightArmCm'];

/** Deterministic index-based generator (no randomness — reproducible timings). */
function synthWorkouts(): TrainingWorkout[] {
  const workouts: TrainingWorkout[] = [];
  for (let d = 0; d < DAYS; d += 1) {
    // ~4 sessions/week: train on 4 of every 7 days.
    if (d % 7 >= 4) continue;
    const date = addDaysIso(TODAY, -(DAYS - d));
    const exercises = Array.from({ length: 5 }, (_, e) => {
      const group = GROUPS[(d + e) % GROUPS.length]!;
      return {
        exerciseId: `ex_${group}_${e}`,
        name: `${group} lift ${e}`,
        loadType: 'external' as const,
        primaryMuscleGroup: group,
        counting: 'none' as const,
        sets: Array.from({ length: 4 }, (_, s) => ({
          weightKg: 40 + ((d + e * 3 + s) % 60),
          reps: 5 + (s % 4),
          warmup: s === 0,
        })),
      };
    });
    workouts.push({ id: `w${d}`, date, startedAt: d * 1000, endedAt: d * 1000 + 3600, exercises });
  }
  return workouts;
}

function synthWeighIns(): WeighIn[] {
  const out: WeighIn[] = [];
  for (let d = 0; d < DAYS; d += 2) {
    out.push({ date: addDaysIso(TODAY, -(DAYS - d)), weightKg: 85 - Math.sin(d / 120) * 4 });
  }
  return out;
}

function synthSnapshots(): BodySnapshot[] {
  const out: BodySnapshot[] = [];
  for (let d = 0; d < DAYS; d += 7) {
    const base = {
      date: addDaysIso(TODAY, -(DAYS - d)),
      weightKg: 85 - Math.sin(d / 120) * 4,
      bodyFatPct: 18 - Math.sin(d / 200) * 3,
      waistCm: 84 - Math.sin(d / 200) * 3,
      chestCm: 100 + Math.sin(d / 200) * 2,
      leftArmCm: 38 + Math.sin(d / 300),
      rightArmCm: 38 + Math.sin(d / 300),
    };
    out.push({
      muscleMassKg: null,
      visceralFat: null,
      bmi: null,
      neckCm: null,
      hipsCm: null,
      leftForearmCm: null,
      rightForearmCm: null,
      leftThighCm: null,
      rightThighCm: null,
      leftCalfCm: null,
      rightCalfCm: null,
      ...base,
    } as BodySnapshot);
  }
  return out;
}

function synthNutrition(): DailyNutrition[] {
  const out: DailyNutrition[] = [];
  for (let d = 0; d < DAYS; d += 1) {
    out.push({
      date: addDaysIso(TODAY, -(DAYS - d)),
      totals: { kcal: 2200 + (d % 400), proteinG: 150 + (d % 60), carbG: 220, fatG: 70 },
      target: { kcal: 2500, proteinG: 180, carbG: 250, fatG: 80, waterMl: 3000 },
      waterMl: 2500 + (d % 800),
      logged: d % 5 !== 0, // ~80% logged
    });
  }
  return out;
}

describe('analytics performance over a 5-year dataset (§7 watchdog)', () => {
  const workouts = synthWorkouts();
  const weighIns = synthWeighIns();
  const snapshots = synthSnapshots();
  const nutritionDays = synthNutrition();
  const window = rangeWindow('90d', TODAY);

  it('reports dataset scale', () => {
    const sets = workouts.reduce(
      (n, w) => n + w.exercises.reduce((m, e) => m + e.sets.length, 0),
      0,
    );
    // Realistic 5-year single-user scale.
    expect(workouts.length).toBeGreaterThan(900);
    expect(sets).toBeGreaterThan(18_000);
    expect(nutritionDays.length).toBeGreaterThan(1800);
    console.log(
      `[perf] 5y dataset: ${workouts.length} workouts, ${sets} sets, ${snapshots.length} snapshots, ${nutritionDays.length} nutrition days`,
    );
  });

  it('computes the full dashboard analytics within a generous ceiling', () => {
    const start = performance.now();
    const workout = computeWorkoutAnalytics({
      workouts,
      weighIns,
      weeklyWorkoutTarget: 4,
      defaultBodyweightKg: 85,
      scheduledWeekdays: [0, 1, 2, 3],
      window,
      today: TODAY,
    });
    const muscle = computeMuscleReports({
      workouts,
      weighIns,
      defaultBodyweightKg: 85,
      window,
      today: TODAY,
    });
    const body = computeBodyAnalytics({
      snapshots,
      window,
      targetWeightKg: 78,
      siteFields: SITE_FIELDS,
    });
    const nutrition = computeNutritionAnalytics({ days: nutritionDays, window, today: TODAY });
    const recomp = computeRecompSignal(snapshots, TODAY);
    const elapsed = performance.now() - start;
    console.log(`[perf] full dashboard analytics over 5y: ${elapsed.toFixed(1)} ms`);
    expect(workout.totalWorkouts).toBeGreaterThan(0);
    expect(muscle.length).toBeGreaterThan(0);
    expect(body.weight.latestKg).not.toBeNull();
    expect(nutrition.loggedDays).toBeGreaterThan(0);
    expect(recomp.status).toBeDefined();
    // Node ceiling far above the 50 ms iPhone budget — a regression guard, not the gate.
    expect(elapsed).toBeLessThan(1500);
  });

  it('buckets a full 5-year chart series within a generous ceiling', () => {
    const series: SeriesPoint[] = snapshots
      .filter((s) => s.weightKg !== null)
      .map((s) => ({ date: s.date, value: s.weightKg! }));
    const start = performance.now();
    const bucketed = bucketSeries(series, 'all', 'mean');
    const elapsed = performance.now() - start;
    console.log(
      `[perf] chart series bucketing (${series.length}→${bucketed.length} pts): ${elapsed.toFixed(2)} ms`,
    );
    expect(bucketed.length).toBeLessThanOrEqual(120);
    expect(elapsed).toBeLessThan(200);
  });
});
