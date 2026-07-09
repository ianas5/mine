/** @jest-environment node */
import { setDbForTesting } from '@/core/db';
import { scalePortion, sumMacros } from '@/domain/nutrition';

import { seedDatabase } from '../seed/seedDatabase';
import { createTestDb, type TestDb } from '../testing/createTestDb';
import { nutritionRepository, type NewMealEntryInput } from './nutritionRepository';

const CHICKEN = 'food_seed_chicken-breast';
const DATE = '2026-07-09';

const logChicken = async (amount: number, slot: NewMealEntryInput['slot'] = 'lunch') => {
  const food = (await nutritionRepository.getFood(CHICKEN))!;
  const m = scalePortion(food, amount);
  return nutritionRepository.addMealEntry({
    date: DATE,
    slot,
    foodId: food.id,
    foodName: food.name,
    loggedAmount: amount,
    loggedUnit: food.servingUnit,
    kcal: m.kcal,
    proteinG: m.proteinG,
    carbG: m.carbG,
    fatG: m.fatG,
  });
};

describe('nutritionRepository (real SQLite)', () => {
  let testDb: TestDb;

  beforeEach(async () => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
    await seedDatabase();
  });

  afterEach(() => testDb.close());

  it('seeds starter foods including quick meals', async () => {
    const foods = await nutritionRepository.listFoods();
    expect(foods.length).toBeGreaterThanOrEqual(15);
    expect(foods.some((f) => f.isQuickMeal)).toBe(true);
  });

  it('snapshots macros so editing a food never rewrites past entries (§4.1)', async () => {
    await logChicken(150); // 165→248 kcal at 1.5×
    const before = (await nutritionRepository.listMealEntries(DATE))[0]!;
    expect(before).toMatchObject({
      kcal: 248,
      proteinG: 46.5,
      foodName: 'Chicken Breast (cooked)',
    });

    // Redefine the food entirely.
    await nutritionRepository.updateFood(CHICKEN, {
      name: 'Chicken Breast (raw)',
      servingAmount: 100,
      servingUnit: 'g',
      kcal: 120,
      proteinG: 22,
      carbG: 0,
      fatG: 2.6,
      isQuickMeal: false,
    });

    const after = (await nutritionRepository.listMealEntries(DATE))[0]!;
    expect(after.kcal).toBe(248);
    expect(after.proteinG).toBe(46.5);
    expect(after.foodName).toBe('Chicken Breast (cooked)');
  });

  it('day totals equal the sum of the day’s entries', async () => {
    await logChicken(100);
    await logChicken(200, 'dinner');
    const entries = await nutritionRepository.listMealEntries(DATE);
    const totals = sumMacros(entries);
    expect(totals.kcal).toBe(165 + 330);
    expect(totals.proteinG).toBeCloseTo(31 + 62, 1);
  });

  it('restores an identical row on undo', async () => {
    const id = await logChicken(150);
    const entry = (await nutritionRepository.listMealEntries(DATE)).find((e) => e.id === id)!;

    await nutritionRepository.deleteMealEntry(id);
    expect(await nutritionRepository.listMealEntries(DATE)).toHaveLength(0);

    await nutritionRepository.restoreMealEntry(entry);
    const restored = (await nutritionRepository.listMealEntries(DATE))[0]!;
    expect(restored).toEqual(entry);
  });

  it('keeps meal history when its food is deleted (SET NULL provenance)', async () => {
    await logChicken(100);
    expect(await nutritionRepository.foodHasHistory(CHICKEN)).toBe(true);

    await nutritionRepository.deleteFood(CHICKEN);

    const entry = (await nutritionRepository.listMealEntries(DATE))[0]!;
    expect(entry.foodId).toBeNull();
    expect(entry.foodName).toBe('Chicken Breast (cooked)'); // snapshot survives
    expect(entry.kcal).toBe(165);
  });

  it('surfaces used foods first with their last-used portion (picker)', async () => {
    await logChicken(175);
    const picks = await nutritionRepository.getFoodPicks();
    const chicken = picks.find((p) => p.food.id === CHICKEN)!;
    expect(chicken.useCount).toBe(1);
    expect(chicken.lastAmount).toBe(175);
    expect(chicken.slot).toBe('lunch');
    // Quick meals are pinned ahead of never-used plain foods.
    const firstPlainUnused = picks.findIndex((p) => !p.food.isQuickMeal && p.useCount === 0);
    const anyQuick = picks.findIndex((p) => p.food.isQuickMeal);
    expect(anyQuick).toBeLessThan(firstPlainUnused);
  });
});
