import type { IsoDate } from '@/core/utils';

import { epleyOneRepMax, isE1rmEligible, isWorkingSet } from './sets';
import type { LoadType } from './taxonomy';

export interface PreviewSet {
  readonly weightKg: number;
  readonly reps: number;
}

/** One historical set of an exercise, flattened with its workout's date/order. */
export interface HistorySetRow {
  readonly workoutId: string;
  readonly date: IsoDate;
  /** Higher = more recent (the workout's created_at). Used to find the last session. */
  readonly workoutOrder: number;
  readonly weightKg: number;
  readonly reps: number;
  readonly warmup: boolean;
}

/** The per-exercise history summary shown at the point of logging (UI_UX §4.1). */
export interface ExercisePreview {
  readonly last: { readonly date: IsoDate; readonly sets: readonly PreviewSet[] } | null;
  readonly bestWeightSet: PreviewSet | null;
  readonly bestE1rmKg: number | null;
}

const EMPTY: ExercisePreview = { last: null, bestWeightSet: null, bestE1rmKg: null };

/**
 * Summarizes an exercise's history into Last / Best / Best e1RM (FITNESS_DOMAIN
 * §3.2 working-set rules, §3.5 e1RM). Pure; returns an all-null preview for a
 * first-ever exercise (no fabricated data, P8).
 */
export function summarizeExerciseHistory(
  rows: readonly HistorySetRow[],
  loadType: LoadType,
): ExercisePreview {
  const working = rows.filter((r) =>
    isWorkingSet({ weightKg: r.weightKg, reps: r.reps, warmup: r.warmup }, loadType),
  );
  if (working.length === 0) {
    return EMPTY;
  }

  const newestOrder = Math.max(...working.map((r) => r.workoutOrder));
  const lastRows = working.filter((r) => r.workoutOrder === newestOrder);
  const last = {
    date: lastRows[0]!.date,
    sets: lastRows.map((r) => ({ weightKg: r.weightKg, reps: r.reps })),
  };

  let bestWeightSet: PreviewSet | null = null;
  let bestE1rmKg: number | null = null;
  for (const r of working) {
    if (bestWeightSet === null || r.weightKg > bestWeightSet.weightKg) {
      bestWeightSet = { weightKg: r.weightKg, reps: r.reps };
    }
    if (r.weightKg > 0 && isE1rmEligible(r.reps)) {
      const e1rm = epleyOneRepMax(r.weightKg, r.reps);
      if (bestE1rmKg === null || e1rm > bestE1rmKg) bestE1rmKg = e1rm;
    }
  }

  return { last, bestWeightSet, bestE1rmKg };
}
