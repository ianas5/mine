import type { IsoDate } from '@/core/utils';

import { BODY_FIELDS, type BodyField } from './fields';

/** The value map of a body record (all metrics optional/nullable). */
export type BodyValues = { readonly [K in BodyField]: number | null };

/** A per-date body record (FITNESS_DOMAIN §5.1). One per calendar date. */
export type BodySnapshot = { readonly date: IsoDate } & BodyValues;

/** The most recent non-null value of a field, with the date it was recorded. */
export interface FieldLatest {
  readonly value: number;
  readonly date: IsoDate;
}

/** BMI from weight + height (FITNESS_DOMAIN §5.2): `weight / heightM²`, or null. */
export function deriveBmi(weightKg: number | null, heightCm: number | null): number | null {
  if (weightKg === null || heightCm === null || heightCm <= 0) return null;
  const heightM = heightCm / 100;
  return weightKg / (heightM * heightM);
}

/** A date's BMI: the entered value wins, else derive from that date's weight (§5.2). */
export function resolveBmi(snapshot: BodySnapshot, heightCm: number | null): number | null {
  return snapshot.bmi ?? deriveBmi(snapshot.weightKg, heightCm);
}

/**
 * The latest non-null value of every field across history (for the current-state
 * display and the Add-Measurements placeholders). `snapshots` must be ordered
 * newest-first. Fields never recorded map to null — no fabricated baseline (P8).
 */
export function latestFieldValues(
  snapshots: readonly BodySnapshot[],
): Record<BodyField, FieldLatest | null> {
  const out = {} as Record<BodyField, FieldLatest | null>;
  for (const field of BODY_FIELDS) {
    let latest: FieldLatest | null = null;
    for (const snapshot of snapshots) {
      const value = snapshot[field];
      if (value !== null) {
        latest = { value, date: snapshot.date };
        break;
      }
    }
    out[field] = latest;
  }
  return out;
}

/**
 * Fields co-logged in ≥ `threshold` of past sessions (UI_UX §5.2 smart default) —
 * these start expanded; the rest sit behind "More sites." With no history, only
 * weight (the anchor metric) is expanded.
 */
export function frequentlyLoggedFields(
  snapshots: readonly BodySnapshot[],
  threshold = 0.5,
): ReadonlySet<BodyField> {
  if (snapshots.length === 0) return new Set<BodyField>(['weightKg']);
  const result = new Set<BodyField>();
  for (const field of BODY_FIELDS) {
    const present = snapshots.reduce((n, s) => (s[field] !== null ? n + 1 : n), 0);
    if (present / snapshots.length >= threshold) result.add(field);
  }
  return result;
}

export interface WeightLogEntry {
  readonly date: IsoDate;
  readonly weightKg: number;
  /** Change from the previous (older) weigh-in, or null for the first. */
  readonly deltaKg: number | null;
}

/**
 * Weigh-ins with the change from the previous weigh-in (Measurements home log).
 * `points` must be newest-first; each delta compares against the next (older) one.
 */
export function weightLogWithDeltas(
  points: readonly { readonly date: IsoDate; readonly weightKg: number }[],
): WeightLogEntry[] {
  return points.map((point, index) => {
    const older = points[index + 1];
    return {
      date: point.date,
      weightKg: point.weightKg,
      deltaKg: older ? Math.round((point.weightKg - older.weightKg) * 10) / 10 : null,
    };
  });
}
