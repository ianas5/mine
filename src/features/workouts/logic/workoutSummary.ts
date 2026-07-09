import { computeWorkoutStats, type WorkoutStats } from '@/domain/fitness';
import type { Workout } from '@/domain/models';

/** Working-set count + volume for a saved workout (FITNESS_DOMAIN §3.5). Pure. */
export function workoutStats(workout: Workout, bodyweightKg: number | null): WorkoutStats {
  return computeWorkoutStats(
    workout.exercises.map((ex) => ({
      loadType: ex.loadType,
      unilateralCounting: ex.unilateralCounting,
      sets: ex.sets.map((s) => ({ weightKg: s.weightKg, reps: s.reps, warmup: s.warmup })),
    })),
    bodyweightKg,
  );
}
