import { E1RM_MAX_TRUSTED_REPS } from './constants';
import type { LoadType } from './taxonomy';

/** How a unilateral exercise entry is counted (FITNESS_DOMAIN §3.4 / edge case 6). */
export type UnilateralCounting = 'none' | 'single_doubled' | 'per_side';

/** The minimal set shape the fitness formulas operate on. */
export interface RawSet {
  readonly weightKg: number;
  readonly reps: number;
  readonly warmup: boolean;
}

/**
 * Effective load in kg by load type (FITNESS_DOMAIN §3.4). `bodyweightKg` is the
 * resolved bodyweight or null when unknown; for `timed` the load is undefined (0).
 */
export function effectiveLoadKg(
  loadType: LoadType,
  weightKg: number,
  bodyweightKg: number | null,
): number {
  switch (loadType) {
    case 'external':
      return weightKg;
    case 'bodyweight':
      return bodyweightKg ?? 0;
    case 'bodyweight_plus':
      return (bodyweightKg ?? 0) + weightKg;
    case 'assisted':
      return Math.max((bodyweightKg ?? 0) - weightKg, 0);
    case 'timed':
      return 0;
  }
}

/**
 * Working-set rule (FITNESS_DOMAIN §3.2): excludes warm-ups and reps<1. External
 * sets require a positive weight; bodyweight-family and timed sets always provide
 * stimulus, so they count even when bodyweight is unknown (edge case 5).
 */
export function isWorkingSet(set: RawSet, loadType: LoadType): boolean {
  if (set.warmup) return false;
  if (set.reps < 1) return false;
  if (loadType === 'external') return set.weightKg > 0;
  return true;
}

/** Set volume load in kg (FITNESS_DOMAIN §3.5); doubled for a single-logged unilateral entry. */
export function setVolumeKg(
  set: RawSet,
  loadType: LoadType,
  bodyweightKg: number | null,
  counting: UnilateralCounting,
): number {
  if (loadType === 'timed' || !isWorkingSet(set, loadType)) {
    return 0;
  }
  const base = effectiveLoadKg(loadType, set.weightKg, bodyweightKg) * set.reps;
  return counting === 'single_doubled' ? base * 2 : base;
}

/**
 * True when a working set's volume is unreliable because bodyweight is unknown
 * for a bodyweight-derived load (FITNESS_DOMAIN §3.4 / edge case 5).
 */
export function isVolumeLowConfidence(
  set: RawSet,
  loadType: LoadType,
  bodyweightKg: number | null,
): boolean {
  if (!isWorkingSet(set, loadType)) return false;
  if (bodyweightKg !== null) return false;
  return loadType === 'bodyweight' || loadType === 'bodyweight_plus' || loadType === 'assisted';
}

/** Epley estimated 1RM (FITNESS_DOMAIN §3.5). Returns 0 for non-positive inputs; w at 1 rep. */
export function epleyOneRepMax(weightKg: number, reps: number): number {
  if (weightKg <= 0 || reps <= 0) return 0;
  if (reps === 1) return weightKg;
  return weightKg * (1 + reps / 30);
}

/** e1RM counts toward PRs only at reps 1..12 (FITNESS_DOMAIN §3.5). */
export function isE1rmEligible(reps: number): boolean {
  return reps >= 1 && reps <= E1RM_MAX_TRUSTED_REPS;
}
