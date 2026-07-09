import { todayIso } from '@/core/utils';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import { detectNewPRs, type ExerciseSetRow, type LoadType, type PrEvent } from '@/domain/fitness';

import type { SessionState } from '../stores/useSessionStore';

export interface ExercisePrGroup {
  readonly exerciseName: string;
  readonly events: readonly PrEvent[];
}

export interface SessionPrs {
  readonly totalCount: number;
  readonly groups: readonly ExercisePrGroup[];
}

export const NO_PRS: SessionPrs = { totalCount: 0, groups: [] };

interface Candidate {
  readonly name: string;
  readonly loadType: LoadType;
  readonly sets: ExerciseSetRow[];
}

/**
 * PR events the current session establishes vs. all prior history (FITNESS_DOMAIN
 * §3.7), computed before the durable write so the finish summary and toast can
 * state "PRs earned". Trustworthy by construction: prior comes from the `sets`
 * table (not a cache) and detection is strictly-greater. Session sets are grouped
 * by exerciseId so the same lift logged twice is judged once against its history.
 */
export async function computeSessionPRs(
  session: SessionState,
  bodyweightKg: number | null,
): Promise<SessionPrs> {
  const candidates = new Map<string, Candidate>();
  const today = todayIso();

  for (const exercise of session.exercises) {
    const candidate: Candidate = candidates.get(exercise.exerciseId) ?? {
      name: exercise.name,
      loadType: exercise.loadType,
      sets: [],
    };
    for (const set of exercise.sets) {
      candidate.sets.push({
        workoutId: 'session',
        date: today,
        // The session is more recent than any saved workout.
        workoutOrder: Number.MAX_SAFE_INTEGER,
        weightKg: set.weightKg,
        reps: set.reps,
        warmup: set.warmup,
        counting: exercise.unilateralCounting,
      });
    }
    candidates.set(exercise.exerciseId, candidate);
  }

  const groups: ExercisePrGroup[] = [];
  for (const [exerciseId, candidate] of candidates) {
    const history = await workoutRepository.getExerciseSetHistory(exerciseId);
    const prior = history?.rows ?? [];
    const events = detectNewPRs(prior, candidate.sets, candidate.loadType, bodyweightKg);
    if (events.length > 0) groups.push({ exerciseName: candidate.name, events });
  }

  return { totalCount: groups.reduce((n, g) => n + g.events.length, 0), groups };
}
