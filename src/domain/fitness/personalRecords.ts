import type { IsoDate } from '@/core/utils';

import { effectiveLoadKg, epleyOneRepMax, isE1rmEligible, isWorkingSet, setVolumeKg } from './sets';
import type { UnilateralCounting } from './sets';
import type { LoadType } from './taxonomy';

/**
 * One historical/candidate set of an exercise, flattened with its workout grouping
 * and per-entry unilateral counting. The single input row shape for both PR
 * detection (FITNESS_DOMAIN §3.7) and the Exercise Report (ANALYTICS §5.5).
 */
export interface ExerciseSetRow {
  readonly workoutId: string;
  readonly date: IsoDate;
  /** The workout's created_at — higher = more recent (finds "last performed"). */
  readonly workoutOrder: number;
  readonly weightKg: number;
  readonly reps: number;
  readonly warmup: boolean;
  readonly counting: UnilateralCounting;
}

/** The scalar all-time bests shown on the report and compared for PR events (§3.7). */
export interface ExerciseBests {
  readonly heaviestWeightKg: number | null;
  readonly bestE1rmKg: number | null;
  readonly bestSetVolumeKg: number | null;
  readonly bestSessionVolumeKg: number | null;
}

export const EMPTY_BESTS: ExerciseBests = {
  heaviestWeightKg: null,
  bestE1rmKg: null,
  bestSetVolumeKg: null,
  bestSessionVolumeKg: null,
};

const raise = (current: number | null, next: number): number =>
  current === null ? next : Math.max(current, next);

const toRawSet = (row: ExerciseSetRow) => ({
  weightKg: row.weightKg,
  reps: row.reps,
  warmup: row.warmup,
});

/**
 * Current all-time bests across an exercise's working sets (FITNESS_DOMAIN §3.7).
 * Pure and recompute-from-history — nothing is cached, so edits and deletes make
 * records recede automatically. Weight and e1RM use single-side `effectiveLoad`
 * (unilateral doubling never inflates a strength PR, §3.4); volume uses the
 * doubled figure where the entry is `single_doubled`.
 */
export function computeExerciseBests(
  rows: readonly ExerciseSetRow[],
  loadType: LoadType,
  bodyweightKg: number | null,
): ExerciseBests {
  let heaviestWeightKg: number | null = null;
  let bestE1rmKg: number | null = null;
  let bestSetVolumeKg: number | null = null;
  const sessionVolume = new Map<string, number>();

  for (const row of rows) {
    if (!isWorkingSet(toRawSet(row), loadType)) continue;

    const eff = effectiveLoadKg(loadType, row.weightKg, bodyweightKg);
    if (eff > 0) {
      heaviestWeightKg = raise(heaviestWeightKg, eff);
      if (isE1rmEligible(row.reps)) bestE1rmKg = raise(bestE1rmKg, epleyOneRepMax(eff, row.reps));
    }

    const volume = setVolumeKg(toRawSet(row), loadType, bodyweightKg, row.counting);
    if (volume > 0) {
      bestSetVolumeKg = raise(bestSetVolumeKg, volume);
      sessionVolume.set(row.workoutId, (sessionVolume.get(row.workoutId) ?? 0) + volume);
    }
  }

  let bestSessionVolumeKg: number | null = null;
  for (const total of sessionVolume.values())
    bestSessionVolumeKg = raise(bestSessionVolumeKg, total);

  return { heaviestWeightKg, bestE1rmKg, bestSetVolumeKg, bestSessionVolumeKg };
}

export type PrKind = 'weight' | 'e1rm' | 'setVolume' | 'sessionVolume' | 'repAtLoad';

/** A newly-set record. `value` is kg for weight/e1RM/volume; reps for `repAtLoad`. */
export interface PrEvent {
  readonly kind: PrKind;
  readonly value: number;
  /** Present only for `repAtLoad`: the effective load (kg) the rep record was set at. */
  readonly loadKg?: number;
}

/** Most reps performed at each distinct effective load among working sets. */
function maxRepsByLoad(
  rows: readonly ExerciseSetRow[],
  loadType: LoadType,
  bodyweightKg: number | null,
): Map<number, number> {
  const byLoad = new Map<number, number>();
  for (const row of rows) {
    if (!isWorkingSet(toRawSet(row), loadType)) continue;
    const eff = effectiveLoadKg(loadType, row.weightKg, bodyweightKg);
    if (eff <= 0) continue;
    byLoad.set(eff, Math.max(byLoad.get(eff) ?? 0, row.reps));
  }
  return byLoad;
}

const beats = (candidate: number | null, prior: number | null): boolean =>
  candidate !== null && (prior === null || candidate > prior);

/**
 * The records `candidate` sets newly establish relative to all `prior` history —
 * strictly greater only, ties are never PRs (FITNESS_DOMAIN §3.7). Warm-ups and
 * loadless/timed sets are excluded by the working-set rule. A rep-at-load PR is
 * only reported when that exact load was lifted before (a first-ever load is
 * already covered by the heaviest-weight PR — no noisy "record" for every new
 * weight). This is the conservative, trustworthy reading of §3.7.
 */
export function detectNewPRs(
  prior: readonly ExerciseSetRow[],
  candidate: readonly ExerciseSetRow[],
  loadType: LoadType,
  bodyweightKg: number | null,
): PrEvent[] {
  const priorBests = computeExerciseBests(prior, loadType, bodyweightKg);
  const candidateBests = computeExerciseBests(candidate, loadType, bodyweightKg);
  const events: PrEvent[] = [];

  if (beats(candidateBests.heaviestWeightKg, priorBests.heaviestWeightKg)) {
    events.push({ kind: 'weight', value: candidateBests.heaviestWeightKg! });
  }
  if (beats(candidateBests.bestE1rmKg, priorBests.bestE1rmKg)) {
    events.push({ kind: 'e1rm', value: candidateBests.bestE1rmKg! });
  }
  if (beats(candidateBests.bestSetVolumeKg, priorBests.bestSetVolumeKg)) {
    events.push({ kind: 'setVolume', value: candidateBests.bestSetVolumeKg! });
  }
  if (beats(candidateBests.bestSessionVolumeKg, priorBests.bestSessionVolumeKg)) {
    events.push({ kind: 'sessionVolume', value: candidateBests.bestSessionVolumeKg! });
  }

  const priorReps = maxRepsByLoad(prior, loadType, bodyweightKg);
  const candidateReps = maxRepsByLoad(candidate, loadType, bodyweightKg);
  for (const [load, reps] of candidateReps) {
    const before = priorReps.get(load);
    if (before !== undefined && reps > before) {
      events.push({ kind: 'repAtLoad', value: reps, loadKg: load });
    }
  }

  return events;
}
