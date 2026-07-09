/** Serving units for foods (DATABASE §3.5, FITNESS_DOMAIN §4.1). */
export const SERVING_UNITS = ['g', 'ml', 'piece', 'scoop', 'cup', 'serving'] as const;
export type ServingUnit = (typeof SERVING_UNITS)[number];

/** Meal slots (FITNESS_DOMAIN §4.1). A meal entry may be untagged (null slot). */
export const MEAL_SLOTS = ['breakfast', 'lunch', 'dinner', 'snacks'] as const;
export type MealSlot = (typeof MEAL_SLOTS)[number];

export const MEAL_SLOT_LABELS: Record<MealSlot, string> = {
  breakfast: 'Breakfast',
  lunch: 'Lunch',
  dinner: 'Dinner',
  snacks: 'Snacks',
};

/** Time-of-day slot fallback when a food has no logged-slot history (UI_UX §5.2). */
export function defaultSlotForHour(hour: number): MealSlot {
  if (hour < 11) return 'breakfast';
  if (hour < 15) return 'lunch';
  if (hour < 21) return 'dinner';
  return 'snacks';
}
