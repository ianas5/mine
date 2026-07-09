/** The four tracked macros (FITNESS_DOMAIN §4). No fiber. `kcal` is an integer. */
export interface MacroSet {
  readonly kcal: number;
  readonly proteinG: number;
  readonly carbG: number;
  readonly fatG: number;
}

export const ZERO_MACROS: MacroSet = { kcal: 0, proteinG: 0, carbG: 0, fatG: 0 };

/** The per-serving definition a portion is scaled from. */
export interface FoodMacros extends MacroSet {
  readonly servingAmount: number;
}

/** Grams to 0.1 g precision (FITNESS_DOMAIN §2.2); avoids float drift. */
const round1 = (value: number): number => Math.round(value * 10) / 10;

/**
 * Portion scaling (FITNESS_DOMAIN §4.2): logged macros = per-serving macros ×
 * (loggedAmount / servingAmount). Computed at log time and stored (snapshot), so a
 * later edit to the food never rewrites the entry. kcal is rounded to an integer,
 * grams to 0.1 g. A non-positive serving amount scales to zero (never divides by 0).
 */
export function scalePortion(food: FoodMacros, loggedAmount: number): MacroSet {
  const factor = food.servingAmount > 0 ? loggedAmount / food.servingAmount : 0;
  return {
    kcal: Math.round(food.kcal * factor),
    proteinG: round1(food.proteinG * factor),
    carbG: round1(food.carbG * factor),
    fatG: round1(food.fatG * factor),
  };
}

/** Day/range totals (FITNESS_DOMAIN §4.2): Σ over entries. Grams re-rounded to 0.1 g. */
export function sumMacros(entries: readonly MacroSet[]): MacroSet {
  const total = entries.reduce<MacroSet>(
    (acc, m) => ({
      kcal: acc.kcal + m.kcal,
      proteinG: acc.proteinG + m.proteinG,
      carbG: acc.carbG + m.carbG,
      fatG: acc.fatG + m.fatG,
    }),
    ZERO_MACROS,
  );
  return {
    kcal: total.kcal,
    proteinG: round1(total.proteinG),
    carbG: round1(total.carbG),
    fatG: round1(total.fatG),
  };
}

/** Atwater energy estimate: 4·protein + 4·carb + 9·fat (FITNESS_DOMAIN §4.2). */
export function atwaterKcal(macros: MacroSet): number {
  return 4 * macros.proteinG + 4 * macros.carbG + 9 * macros.fatG;
}

/**
 * Macro/energy cross-check (FITNESS_DOMAIN §4.2): flags a food whose entered kcal
 * diverges from its Atwater estimate by more than `tolerance` (default 20 %). A
 * *validation aid only* — it never overrides the entered kcal. Foods with 0 kcal
 * are not flagged (nothing to compare against).
 */
export function isKcalImplausible(macros: MacroSet, tolerance = 0.2): boolean {
  if (macros.kcal <= 0) return false;
  return Math.abs(macros.kcal - atwaterKcal(macros)) / macros.kcal > tolerance;
}
