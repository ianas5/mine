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
