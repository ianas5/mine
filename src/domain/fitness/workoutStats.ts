import {
  isVolumeLowConfidence,
  isWorkingSet,
  setVolumeKg,
  type RawSet,
  type UnilateralCounting,
} from './sets';
import type { LoadType } from './taxonomy';

export interface StatExercise {
  readonly loadType: LoadType;
  readonly unilateralCounting: UnilateralCounting;
  readonly sets: readonly RawSet[];
}

export interface WorkoutStats {
  readonly workingSetCount: number;
  readonly totalVolumeKg: number;
  /** At least one working set has unreliable volume (unknown bodyweight). */
  readonly volumeLowConfidence: boolean;
}

/** Aggregates working sets and volume for a session/workout (FITNESS_DOMAIN §3.5). */
export function computeWorkoutStats(
  exercises: readonly StatExercise[],
  bodyweightKg: number | null,
): WorkoutStats {
  let workingSetCount = 0;
  let totalVolumeKg = 0;
  let volumeLowConfidence = false;

  for (const exercise of exercises) {
    for (const set of exercise.sets) {
      if (!isWorkingSet(set, exercise.loadType)) continue;
      workingSetCount += 1;
      totalVolumeKg += setVolumeKg(
        set,
        exercise.loadType,
        bodyweightKg,
        exercise.unilateralCounting,
      );
      if (isVolumeLowConfidence(set, exercise.loadType, bodyweightKg)) {
        volumeLowConfidence = true;
      }
    }
  }

  return { workingSetCount, totalVolumeKg, volumeLowConfidence };
}

/** A workout counts toward history/consistency only with ≥1 working set (FITNESS_DOMAIN §3.8, edge 2). */
export function isCountableWorkout(
  exercises: readonly StatExercise[],
  bodyweightKg: number | null,
): boolean {
  return computeWorkoutStats(exercises, bodyweightKg).workingSetCount > 0;
}
