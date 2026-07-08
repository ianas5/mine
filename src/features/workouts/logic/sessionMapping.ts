import type { StatExercise } from '@/domain/fitness';
import type { NewWorkoutInput } from '@/data/repositories/workoutRepository';

import type { SessionExercise, SessionState } from '../stores/useSessionStore';

/** Session exercises → the shape the stats engine consumes (FITNESS_DOMAIN §3.5). Pure. */
export function sessionToStatExercises(exercises: readonly SessionExercise[]): StatExercise[] {
  return exercises.map((ex) => ({
    loadType: ex.loadType,
    unilateralCounting: ex.unilateralCounting,
    sets: ex.sets.map((s) => ({ weightKg: s.weightKg, reps: s.reps, warmup: s.warmup })),
  }));
}

/** Session → persistence input (DATABASE §7). `endedAt` is passed in to stay pure. */
export function sessionToWorkoutInput(state: SessionState, endedAt: number): NewWorkoutInput {
  return {
    name: state.name,
    startedAt: state.startedAt,
    endedAt,
    notes: null,
    exercises: state.exercises.map((ex) => ({
      exerciseId: ex.exerciseId,
      unilateralCounting: ex.unilateralCounting,
      notes: null,
      sets: ex.sets.map((s) => ({
        weightKg: s.weightKg,
        reps: s.reps,
        rpe: s.rpe,
        warmup: s.warmup,
      })),
    })),
  };
}
