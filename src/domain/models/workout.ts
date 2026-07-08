import type { EpochMs, IsoDate } from '@/core/utils';
import type { LoadType, UnilateralCounting } from '@/domain/fitness';

export interface WorkoutSet {
  readonly id: string;
  readonly position: number;
  readonly weightKg: number;
  readonly reps: number;
  readonly rpe: number | null;
  readonly rir: number | null;
  readonly warmup: boolean;
  readonly notes: string | null;
}

export interface WorkoutExercise {
  readonly id: string;
  readonly exerciseId: string;
  readonly name: string;
  readonly loadType: LoadType;
  readonly unilateralCounting: UnilateralCounting;
  readonly position: number;
  readonly sets: readonly WorkoutSet[];
  readonly notes: string | null;
}

/** A performed training session (FITNESS_DOMAIN §3.1, DATABASE §3.4). */
export interface Workout {
  readonly id: string;
  readonly date: IsoDate;
  readonly name: string;
  readonly startedAt: EpochMs | null;
  readonly endedAt: EpochMs | null;
  readonly notes: string | null;
  readonly exercises: readonly WorkoutExercise[];
}
