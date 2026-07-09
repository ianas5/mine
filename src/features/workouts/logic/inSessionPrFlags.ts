import {
  effectiveLoadKg,
  epleyOneRepMax,
  isE1rmEligible,
  isWorkingSet,
  type ExerciseBests,
  type LoadType,
} from '@/domain/fitness';

import type { SessionSet } from '../stores/useSessionStore';

/**
 * Optimistic in-session PR flags (UI_UX §4.1, P15): marks each set that sets a
 * new *running* weight or e1RM best versus the prior-history baseline plus earlier
 * sets this session — so only genuine record-setters light up (not every set of a
 * first-ever exercise). A hint, never authoritative; the finish summary recomputes
 * the real records from history. "First ever" counts as a PR (delight registry).
 */
export function inSessionPrFlags(
  sets: readonly SessionSet[],
  loadType: LoadType,
  bodyweightKg: number | null,
  baseline: ExerciseBests,
): boolean[] {
  let bestWeight = baseline.heaviestWeightKg;
  let bestE1rm = baseline.bestE1rmKg;

  return sets.map((set) => {
    if (!isWorkingSet({ weightKg: set.weightKg, reps: set.reps, warmup: set.warmup }, loadType)) {
      return false;
    }
    const eff = effectiveLoadKg(loadType, set.weightKg, bodyweightKg);
    if (eff <= 0) return false;

    let isPr = false;
    if (bestWeight === null || eff > bestWeight) {
      isPr = true;
      bestWeight = eff;
    }
    if (isE1rmEligible(set.reps)) {
      const e1rm = epleyOneRepMax(eff, set.reps);
      if (bestE1rm === null || e1rm > bestE1rm) {
        isPr = true;
        bestE1rm = e1rm;
      }
    }
    return isPr;
  });
}
