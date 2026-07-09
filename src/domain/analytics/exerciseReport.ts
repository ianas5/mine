import type { IsoDate } from '@/core/utils';
import {
  computeExerciseBests,
  effectiveLoadKg,
  isWorkingSet,
  setVolumeKg,
  type ExerciseBests,
  type ExerciseSetRow,
  type LoadType,
} from '@/domain/fitness';

import { computeTrend, type Trend } from './trend';
import type { MetricResult } from './metricResult';
import { sortByDate, type SeriesPoint } from './timeSeries';

/**
 * Stability deadband (kg) for the e1RM strength trend. FITNESS_DOMAIN §6.4 fixes
 * deadbands for body metrics only; strength (§3.5/§5.1) has none, so this is an
 * analytics-level constant (a new metric's threshold, not a redefinition of §6.4).
 */
export const E1RM_STABILITY_KG = 2.5;

/**
 * The Exercise Report data contract (ANALYTICS §5.5), all-time. Trend and
 * progression rate are time-series analytics deferred to Phase 15 (stated
 * on-screen, never faked — P8), so they are deliberately absent here. Pure and
 * recompute-from-history, so every number reconciles exactly with the workout
 * detail screen and recedes on edit/delete.
 */
export interface ExerciseReport {
  readonly totalSessions: number;
  readonly totalWorkingSets: number;
  readonly totalVolumeKg: number;
  readonly bests: ExerciseBests;
  readonly avgRepsPerWorkingSet: number | null;
  readonly avgEffectiveLoadKg: number | null;
  readonly lastPerformed: { readonly date: IsoDate; readonly workoutOrder: number } | null;
}

export const EMPTY_REPORT: ExerciseReport = {
  totalSessions: 0,
  totalWorkingSets: 0,
  totalVolumeKg: 0,
  bests: {
    heaviestWeightKg: null,
    bestE1rmKg: null,
    bestSetVolumeKg: null,
    bestSessionVolumeKg: null,
  },
  avgRepsPerWorkingSet: null,
  avgEffectiveLoadKg: null,
  lastPerformed: null,
};

const rawSet = (row: ExerciseSetRow) => ({
  weightKg: row.weightKg,
  reps: row.reps,
  warmup: row.warmup,
});

/** Computes the all-time Exercise Report from an exercise's full set history. */
export function computeExerciseReport(
  rows: readonly ExerciseSetRow[],
  loadType: LoadType,
  bodyweightKg: number | null,
): ExerciseReport {
  const working = rows.filter((row) => isWorkingSet(rawSet(row), loadType));
  if (working.length === 0) return EMPTY_REPORT;

  const sessions = new Set<string>();
  let totalVolumeKg = 0;
  let repsSum = 0;
  let loadSum = 0;
  let loadCount = 0;
  let lastPerformed: ExerciseReport['lastPerformed'] = null;

  for (const row of working) {
    sessions.add(row.workoutId);
    repsSum += row.reps;
    totalVolumeKg += setVolumeKg(rawSet(row), loadType, bodyweightKg, row.counting);

    const eff = effectiveLoadKg(loadType, row.weightKg, bodyweightKg);
    if (eff > 0) {
      loadSum += eff;
      loadCount += 1;
    }

    if (lastPerformed === null || row.workoutOrder > lastPerformed.workoutOrder) {
      lastPerformed = { date: row.date, workoutOrder: row.workoutOrder };
    }
  }

  return {
    totalSessions: sessions.size,
    totalWorkingSets: working.length,
    totalVolumeKg,
    bests: computeExerciseBests(rows, loadType, bodyweightKg),
    avgRepsPerWorkingSet: repsSum / working.length,
    avgEffectiveLoadKg: loadCount > 0 ? loadSum / loadCount : null,
    lastPerformed,
  };
}

/**
 * The per-workout best-e1RM series (ANALYTICS §5.1 "strength trend"): one point per
 * workout = the best e1RM among its e1RM-eligible working sets. Reuses
 * `computeExerciseBests` per workout so it matches the report's `bestE1rmKg` exactly.
 * Ascending by date; workouts with no eligible set contribute no point.
 */
export function bestE1rmSeries(
  rows: readonly ExerciseSetRow[],
  loadType: LoadType,
  bodyweightKg: number | null,
): SeriesPoint[] {
  const byWorkout = new Map<string, { date: IsoDate; rows: ExerciseSetRow[] }>();
  for (const row of rows) {
    const bucket = byWorkout.get(row.workoutId);
    if (bucket) bucket.rows.push(row);
    else byWorkout.set(row.workoutId, { date: row.date, rows: [row] });
  }

  const series: SeriesPoint[] = [];
  for (const { date, rows: workoutRows } of byWorkout.values()) {
    const best = computeExerciseBests(workoutRows, loadType, bodyweightKg).bestE1rmKg;
    if (best !== null) series.push({ date, value: best });
  }
  return sortByDate(series);
}

export interface ExerciseTrend {
  /** Per-workout best e1RM, ascending — the sparkline series. */
  readonly series: SeriesPoint[];
  /** Regression trend; `slopePerWeek` is the progression rate (kg/week) when `ok`. */
  readonly trend: MetricResult<Trend>;
}

/**
 * The e1RM strength trend for an exercise (closes the Phase 7 trend debt). Higher e1RM
 * is better (§5.3); below the §6.4 minimums it returns `insufficient-data` — the report
 * shows the series' latest with no fabricated slope.
 */
export function computeExerciseTrend(
  rows: readonly ExerciseSetRow[],
  loadType: LoadType,
  bodyweightKg: number | null,
  today: IsoDate,
): ExerciseTrend {
  const series = bestE1rmSeries(rows, loadType, bodyweightKg);
  const trend = computeTrend(
    series,
    { stabilityThreshold: E1RM_STABILITY_KG, goodDirection: 'higher', pointNoun: 'sessions' },
    { key: 'all', startDate: series[0]?.date ?? null, endDate: today, days: null },
  );
  return { series, trend };
}
