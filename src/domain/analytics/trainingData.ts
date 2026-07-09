import type { IsoDate } from '@/core/utils';
import type { LoadType, MuscleGroup, UnilateralCounting } from '@/domain/fitness';

/**
 * The normalized training input shared by the Workout and Muscle calculators. The
 * repository joins workouts → exercises → sets and hands over these plain rows; all
 * domain math (working sets, effective load, unilateral doubling, volume, e1RM,
 * bodyweight resolution) happens here in `domain/`, never in SQL (ANALYTICS rule 9).
 */

export interface TrainingSet {
  readonly weightKg: number;
  readonly reps: number;
  readonly warmup: boolean;
}

export interface TrainingExercise {
  readonly exerciseId: string;
  readonly name: string;
  readonly loadType: LoadType;
  readonly primaryMuscleGroup: MuscleGroup;
  readonly counting: UnilateralCounting;
  readonly sets: readonly TrainingSet[];
}

export interface TrainingWorkout {
  readonly id: string;
  readonly date: IsoDate;
  readonly startedAt: number | null;
  readonly endedAt: number | null;
  readonly exercises: readonly TrainingExercise[];
}

export interface WeighIn {
  readonly date: IsoDate;
  readonly weightKg: number;
}

/**
 * §3.4 bodyweight resolution: the most recent weigh-in **on or before** the given date,
 * else the settings fallback, else null (the set still counts for stimulus but adds 0
 * volume, flagged low-confidence). `weighIns` may be in any order.
 */
export function resolveBodyweightForDate(
  weighIns: readonly WeighIn[],
  date: IsoDate,
  fallback: number | null,
): number | null {
  let best: WeighIn | null = null;
  for (const w of weighIns) {
    if (w.date <= date && (best === null || w.date > best.date)) best = w;
  }
  return best?.weightKg ?? fallback;
}
