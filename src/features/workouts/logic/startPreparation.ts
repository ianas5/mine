import { isWorkingSet, type PreviewSet } from '@/domain/fitness';
import type { Template, Workout } from '@/domain/models';
import { workoutRepository } from '@/data/repositories/workoutRepository';

import type { PreparedExercise } from '../stores/useSessionStore';

export interface PreparedStart {
  readonly name: string;
  readonly prepared: PreparedExercise[];
}

/** Pads last-time sets up to `count`, repeating the final set's values (never fabricating reps). */
function padSets(last: readonly PreviewSet[], count: number): PreviewSet[] {
  if (count <= 0) return [...last];
  const out: PreviewSet[] = [];
  for (let i = 0; i < count; i += 1) {
    out.push(last[i] ?? last[last.length - 1] ?? { weightKg: 0, reps: 0 });
  }
  return out;
}

/**
 * Builds a session from a template (Phase 8): each exercise carries its planned
 * target (shown at logging) and is pre-filled from last time's working sets where
 * available, else `target_sets` empty rows. The template is a plan — this reads it,
 * it never writes back to it.
 */
export async function prepareTemplateStart(template: Template): Promise<PreparedStart> {
  const prepared: PreparedExercise[] = [];
  for (const te of template.exercises) {
    const preview = await workoutRepository.getExercisePreview(te.exerciseId, te.loadType);
    const last = preview.last?.sets ?? [];
    const targetSets = te.target.sets ?? 0;

    let sets: PreviewSet[];
    if (last.length > 0) {
      sets = padSets(last, Math.max(last.length, targetSets));
    } else if (targetSets > 0) {
      sets = Array.from({ length: targetSets }, () => ({ weightKg: 0, reps: 0 }));
    } else {
      sets = [];
    }

    prepared.push({
      exerciseId: te.exerciseId,
      name: te.name,
      loadType: te.loadType,
      unilateralCounting: te.defaultUnilateral ? 'single_doubled' : 'none',
      target: {
        sets: te.target.sets,
        repMin: te.target.repMin,
        repMax: te.target.repMax,
        rpe: te.target.rpe,
      },
      restSeconds: te.target.restSeconds,
      sets,
    });
  }
  return { name: template.name, prepared };
}

/**
 * Builds a session that repeats a past workout (UI_UX §5.2): reload its working
 * sets as pre-fill. A convenience over history — the original workout is untouched.
 */
export function prepareRepeatLast(workout: Workout): PreparedStart {
  const prepared: PreparedExercise[] = workout.exercises.map((ex) => ({
    exerciseId: ex.exerciseId,
    name: ex.name,
    loadType: ex.loadType,
    unilateralCounting: ex.unilateralCounting,
    target: null,
    restSeconds: null,
    sets: ex.sets
      .filter((s) =>
        isWorkingSet({ weightKg: s.weightKg, reps: s.reps, warmup: s.warmup }, ex.loadType),
      )
      .map((s) => ({ weightKg: s.weightKg, reps: s.reps })),
  }));
  return { name: workout.name, prepared };
}
