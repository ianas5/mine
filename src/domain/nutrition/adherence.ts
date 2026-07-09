import type { MacroSet } from './macros';

/** The macro figures needed to judge a day (a resolved target satisfies this). */
export interface TargetMacros {
  readonly kcal: number;
  readonly proteinG: number;
  readonly carbG: number;
  readonly fatG: number;
  readonly waterMl: number | null;
}

/**
 * Per-macro adherence outcome (FITNESS_DOMAIN §4.3). `near` is protein-only (the
 * ≥ 90 % soft band under its floor); `over` never applies to protein (more is fine).
 */
export type MacroStatus = 'hit' | 'near' | 'under' | 'over';

export interface DayAdherence {
  readonly calories: MacroStatus;
  readonly protein: MacroStatus;
  readonly carbs: MacroStatus;
  readonly fat: MacroStatus;
  /** null when the active target sets no water goal. */
  readonly water: MacroStatus | null;
}

const round1 = (value: number): number => Math.round(value * 10) / 10;

/** A ±tolerance band around a target (calories ±10 %, carbs/fat ±15 %; §4.3). */
function bandStatus(consumed: number, target: number, tolerance: number): MacroStatus {
  if (target <= 0) return 'hit';
  if (consumed < target * (1 - tolerance)) return 'under';
  if (consumed > target * (1 + tolerance)) return 'over';
  return 'hit';
}

/** A floor with a soft near-band (protein: ≥ target hit, ≥ 90 % near; §4.3). */
function floorStatus(consumed: number, target: number): MacroStatus {
  if (target <= 0) return 'hit';
  if (consumed >= target) return 'hit';
  if (consumed >= target * 0.9) return 'near';
  return 'under';
}

/**
 * A day's adherence against the target that was active on that date (FITNESS_DOMAIN
 * §4.3). The caller resolves the target through the single canonical path; this is
 * pure judgement over the resolved numbers. Water is judged only when the target
 * sets a water goal.
 */
export function dayAdherence(
  target: TargetMacros,
  consumed: MacroSet,
  waterMl: number,
): DayAdherence {
  return {
    calories: bandStatus(consumed.kcal, target.kcal, 0.1),
    protein: floorStatus(consumed.proteinG, target.proteinG),
    carbs: bandStatus(consumed.carbG, target.carbG, 0.15),
    fat: bandStatus(consumed.fatG, target.fatG, 0.15),
    water: target.waterMl !== null ? (waterMl >= target.waterMl ? 'hit' : 'under') : null,
  };
}

/** Remaining per macro = target − consumed (negative = over; FITNESS_DOMAIN §4.2). */
export function remainingMacros(target: TargetMacros, consumed: MacroSet): MacroSet {
  return {
    kcal: target.kcal - consumed.kcal,
    proteinG: round1(target.proteinG - consumed.proteinG),
    carbG: round1(target.carbG - consumed.carbG),
    fatG: round1(target.fatG - consumed.fatG),
  };
}
