import type { ServingUnit } from '@/domain/nutrition';

/** A seed food definition; stable id is `food_seed_<slug>` (DATABASE §5.6). */
export interface SeedFood {
  readonly slug: string;
  readonly name: string;
  readonly servingAmount: number;
  readonly servingUnit: ServingUnit;
  readonly kcal: number;
  readonly proteinG: number;
  readonly carbG: number;
  readonly fatG: number;
  readonly quickMeal?: boolean;
}

const F = (
  slug: string,
  name: string,
  servingAmount: number,
  servingUnit: ServingUnit,
  kcal: number,
  proteinG: number,
  carbG: number,
  fatG: number,
  quickMeal = false,
): SeedFood => ({ slug, name, servingAmount, servingUnit, kcal, proteinG, carbG, fatG, quickMeal });

/**
 * A small, real-world starter set (FITNESS_DOMAIN §4.1). The Recent list absorbs
 * the user's own foods over time; this only bootstraps the first logs. No fiber.
 */
export const SEED_FOODS: readonly SeedFood[] = [
  F('chicken-breast', 'Chicken Breast (cooked)', 100, 'g', 165, 31, 0, 3.6),
  F('white-rice', 'White Rice (cooked)', 100, 'g', 130, 2.7, 28, 0.3),
  F('rolled-oats', 'Rolled Oats (dry)', 40, 'g', 156, 6.8, 26.4, 2.8),
  F('whole-egg', 'Whole Egg', 1, 'piece', 78, 6.3, 0.6, 5.3),
  F('greek-yogurt', 'Greek Yogurt (nonfat)', 100, 'g', 59, 10, 3.6, 0.4),
  F('whey-protein', 'Whey Protein', 1, 'scoop', 120, 24, 3, 1.5),
  F('banana', 'Banana', 1, 'piece', 105, 1.3, 27, 0.4),
  F('apple', 'Apple', 1, 'piece', 95, 0.5, 25, 0.3),
  F('almonds', 'Almonds', 30, 'g', 174, 6.3, 6.6, 15),
  F('peanut-butter', 'Peanut Butter', 32, 'g', 190, 8, 7, 16),
  F('olive-oil', 'Olive Oil', 15, 'g', 133, 0, 0, 15),
  F('broccoli', 'Broccoli', 100, 'g', 34, 2.8, 7, 0.4),
  F('ground-beef-90', 'Ground Beef (90/10, cooked)', 100, 'g', 176, 20, 0, 10),
  F('salmon', 'Salmon (cooked)', 100, 'g', 208, 20, 0, 13),
  F('sweet-potato', 'Sweet Potato (cooked)', 100, 'g', 86, 1.6, 20, 0.1),
  F('milk-2pct', 'Milk (2%)', 1, 'cup', 122, 8, 12, 4.8),
  F('chicken-and-rice', 'Chicken and Rice', 1, 'serving', 400, 40, 45, 8, true),
  F('protein-shake', 'Protein Shake', 1, 'serving', 250, 30, 20, 5, true),
];
