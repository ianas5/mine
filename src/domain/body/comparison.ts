import { BODY_FIELDS, type BodyField } from './fields';
import { resolveBmi, type BodySnapshot } from './snapshot';

/** Which direction is "improving" for a metric (FITNESS_DOMAIN §5.3). */
export type BodyDirection = 'lower' | 'higher' | 'neutral';

export const BODY_DIRECTION: Record<BodyField, BodyDirection> = {
  // Weight and BMI depend on the user's goal/target — no fixed good direction.
  weightKg: 'neutral',
  bmi: 'neutral',
  // Fat/belly sites: lower is better.
  bodyFatPct: 'lower',
  visceralFat: 'lower',
  waistCm: 'lower',
  hipsCm: 'lower',
  // Muscular sites and muscle mass: higher is better.
  muscleMassKg: 'higher',
  neckCm: 'higher',
  chestCm: 'higher',
  leftArmCm: 'higher',
  rightArmCm: 'higher',
  leftForearmCm: 'higher',
  rightForearmCm: 'higher',
  leftThighCm: 'higher',
  rightThighCm: 'higher',
  leftCalfCm: 'higher',
  rightCalfCm: 'higher',
};

/**
 * Stability deadband per metric (FITNESS_DOMAIN §6.4): a change smaller than this
 * is "stable", not a trend. The doc fixes weight/circumference/body-fat/muscle;
 * visceral-fat and BMI use conservative domain defaults in the same spirit.
 */
export const BODY_STABILITY: Record<BodyField, number> = {
  weightKg: 0.8,
  bodyFatPct: 0.4,
  muscleMassKg: 0.3,
  visceralFat: 0.5,
  bmi: 0.3,
  neckCm: 0.5,
  chestCm: 0.5,
  waistCm: 0.5,
  hipsCm: 0.5,
  leftArmCm: 0.5,
  rightArmCm: 0.5,
  leftForearmCm: 0.5,
  rightForearmCm: 0.5,
  leftThighCm: 0.5,
  rightThighCm: 0.5,
  leftCalfCm: 0.5,
  rightCalfCm: 0.5,
};

export type ChangeDirection = 'improving' | 'declining' | 'stable' | 'neutral' | 'incomparable';

export interface FieldComparison {
  readonly field: BodyField;
  readonly a: number | null;
  readonly b: number | null;
  /** B − A, or null when the field is not present on both dates (no fabricated baseline). */
  readonly deltaAbs: number | null;
  /** (B − A) / A × 100, or null when incomparable or A = 0 (§5.4). */
  readonly deltaPct: number | null;
  readonly direction: ChangeDirection;
}

/**
 * The best value recorded for each directional field across history (delight registry —
 * "measurement best"). Best = lowest for `lower` fields, highest for `higher`; neutral
 * fields (weight, BMI — no fixed good direction) are excluded. Pure over snapshots.
 */
export function bestFieldValues(
  snapshots: readonly BodySnapshot[],
): Partial<Record<BodyField, number>> {
  const best: Partial<Record<BodyField, number>> = {};
  for (const snapshot of snapshots) {
    for (const field of BODY_FIELDS) {
      const dir = BODY_DIRECTION[field];
      if (dir === 'neutral') continue;
      const value = snapshot[field];
      if (value === null) continue;
      const current = best[field];
      if (current === undefined || (dir === 'lower' ? value < current : value > current)) {
        best[field] = value;
      }
    }
  }
  return best;
}

/**
 * Whether `value` is a new best for `field` in its good direction (FITNESS_DOMAIN §5.3) —
 * strictly beating the prior best. A field with no prior best, or a neutral-direction field
 * (weight/BMI), is never a "best yet" (there is nothing to beat, or no direction to beat in).
 */
export function isFieldBest(field: BodyField, priorBest: number | null, value: number): boolean {
  const dir = BODY_DIRECTION[field];
  if (dir === 'neutral' || priorBest === null) return false;
  return dir === 'lower' ? value < priorBest : value > priorBest;
}

const round1 = (value: number): number => Math.round(value * 10) / 10;

/** A field's value for comparison; BMI resolves entered-else-derived (§5.2). */
function valueOf(snapshot: BodySnapshot, field: BodyField, heightCm: number | null): number | null {
  return field === 'bmi' ? resolveBmi(snapshot, heightCm) : snapshot[field];
}

/**
 * Compares two body snapshots field by field (FITNESS_DOMAIN §5.4). A field present
 * on only one date is **incomparable** — no change is shown, never a fabricated
 * baseline (P8). For fields present on both: absolute + percentage change, and a
 * direction judged by the §6.4 stability deadband mapped through §5.3 directionality.
 */
export function compareSnapshots(
  a: BodySnapshot,
  b: BodySnapshot,
  heightCm: number | null = null,
): FieldComparison[] {
  return BODY_FIELDS.map((field) => {
    const aVal = valueOf(a, field, heightCm);
    const bVal = valueOf(b, field, heightCm);

    if (aVal === null || bVal === null) {
      return { field, a: aVal, b: bVal, deltaAbs: null, deltaPct: null, direction: 'incomparable' };
    }

    const delta = bVal - aVal;
    const deltaPct = aVal !== 0 ? round1((delta / aVal) * 100) : null;

    let direction: ChangeDirection;
    if (Math.abs(delta) < BODY_STABILITY[field]) {
      direction = 'stable';
    } else if (BODY_DIRECTION[field] === 'neutral') {
      direction = 'neutral';
    } else if (BODY_DIRECTION[field] === 'lower') {
      direction = delta < 0 ? 'improving' : 'declining';
    } else {
      direction = delta > 0 ? 'improving' : 'declining';
    }

    return { field, a: aVal, b: bVal, deltaAbs: round1(delta), deltaPct, direction };
  });
}
