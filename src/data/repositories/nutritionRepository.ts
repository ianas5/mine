import { asc, desc, eq } from 'drizzle-orm';

import { emitTableChanges, getDb } from '@/core/db';
import type { MealSlot, ServingUnit } from '@/domain/nutrition';
import type { Food, MealEntry } from '@/domain/models';

import { newId } from '../id';
import { foods, mealEntries } from '../schema/tables';

export interface FoodInput {
  readonly name: string;
  readonly servingAmount: number;
  readonly servingUnit: ServingUnit;
  readonly kcal: number;
  readonly proteinG: number;
  readonly carbG: number;
  readonly fatG: number;
  readonly isQuickMeal: boolean;
}

export interface NewMealEntryInput {
  readonly date: string;
  readonly slot: MealSlot | null;
  readonly foodId: string | null;
  readonly foodName: string;
  readonly loggedAmount: number;
  readonly loggedUnit: ServingUnit;
  readonly kcal: number;
  readonly proteinG: number;
  readonly carbG: number;
  readonly fatG: number;
}

/** A food surfaced in the Log Meal picker with its smart-default portion + slot. */
export interface FoodPick {
  readonly food: Food;
  readonly lastAmount: number | null;
  readonly slot: MealSlot | null;
  readonly useCount: number;
}

const PICKER_WINDOW = 400; // recent entries scanned for frequency/last-used (UI_UX §4.3)

function rowToFood(r: typeof foods.$inferSelect): Food {
  return {
    id: r.id,
    name: r.name,
    servingAmount: r.servingAmount,
    servingUnit: r.servingUnit as ServingUnit,
    kcal: r.kcal,
    proteinG: r.proteinG,
    carbG: r.carbG,
    fatG: r.fatG,
    isQuickMeal: r.isQuickMeal === 1,
    isCustom: r.isCustom === 1,
    isArchived: r.isArchived === 1,
  };
}

function rowToEntry(r: typeof mealEntries.$inferSelect): MealEntry {
  return {
    id: r.id,
    date: r.date,
    slot: r.slot as MealSlot | null,
    foodId: r.foodId,
    foodName: r.foodName,
    loggedAmount: r.loggedAmount,
    loggedUnit: r.loggedUnit as ServingUnit,
    kcal: r.kcal,
    proteinG: r.proteinG,
    carbG: r.carbG,
    fatG: r.fatG,
    loggedAt: r.loggedAt,
  };
}

const mostFrequent = (slots: readonly (MealSlot | null)[]): MealSlot | null => {
  const counts = new Map<MealSlot, number>();
  for (const s of slots) if (s !== null) counts.set(s, (counts.get(s) ?? 0) + 1);
  let best: MealSlot | null = null;
  let bestCount = 0;
  for (const [slot, count] of counts) {
    if (count > bestCount) {
      bestCount = count;
      best = slot;
    }
  }
  return best;
};

export const nutritionRepository = {
  /** Non-archived foods, alphabetical (the base catalog behind search). */
  async listFoods(): Promise<Food[]> {
    const rows = await getDb()
      .select()
      .from(foods)
      .where(eq(foods.isArchived, 0))
      .orderBy(asc(foods.name));
    return rows.map(rowToFood);
  },

  async getFood(id: string): Promise<Food | null> {
    const rows = await getDb().select().from(foods).where(eq(foods.id, id));
    const row = rows[0];
    return row ? rowToFood(row) : null;
  },

  /**
   * Foods for the Log Meal sheet (UI_UX §4.3): most-used first with quick meals
   * pinned, each carrying its last-used portion and most-frequent slot as smart
   * defaults. Frequency/last-used are computed over a recent window of entries.
   */
  async getFoodPicks(): Promise<FoodPick[]> {
    const db = getDb();
    const foodList = await this.listFoods();
    const recent = await db
      .select({
        foodId: mealEntries.foodId,
        loggedAmount: mealEntries.loggedAmount,
        slot: mealEntries.slot,
        loggedAt: mealEntries.loggedAt,
      })
      .from(mealEntries)
      .orderBy(desc(mealEntries.loggedAt))
      .limit(PICKER_WINDOW);

    interface Usage {
      count: number;
      lastAmount: number | null;
      lastLoggedAt: number;
      slots: (MealSlot | null)[];
    }
    const usage = new Map<string, Usage>();
    for (const r of recent) {
      if (r.foodId === null) continue;
      const u = usage.get(r.foodId);
      if (u) {
        u.count += 1;
        u.slots.push(r.slot as MealSlot | null);
      } else {
        // First seen = most recent (rows are newest-first) → its amount is "last used".
        usage.set(r.foodId, {
          count: 1,
          lastAmount: r.loggedAmount,
          lastLoggedAt: r.loggedAt,
          slots: [r.slot as MealSlot | null],
        });
      }
    }

    return foodList
      .map((food) => {
        const u = usage.get(food.id);
        return {
          food,
          lastAmount: u?.lastAmount ?? null,
          slot: u ? mostFrequent(u.slots) : null,
          useCount: u?.count ?? 0,
        };
      })
      .sort((a, b) => {
        if (a.food.isQuickMeal !== b.food.isQuickMeal) return a.food.isQuickMeal ? -1 : 1;
        if (a.useCount !== b.useCount) return b.useCount - a.useCount;
        const aLast = usage.get(a.food.id)?.lastLoggedAt ?? 0;
        const bLast = usage.get(b.food.id)?.lastLoggedAt ?? 0;
        if (aLast !== bLast) return bLast - aLast;
        return a.food.name.localeCompare(b.food.name);
      });
  },

  async createFood(input: FoodInput): Promise<string> {
    const id = newId('food');
    const now = Date.now();
    await getDb()
      .insert(foods)
      .values({
        id,
        name: input.name.trim() || 'Food',
        servingAmount: input.servingAmount,
        servingUnit: input.servingUnit,
        kcal: input.kcal,
        proteinG: input.proteinG,
        carbG: input.carbG,
        fatG: input.fatG,
        isQuickMeal: input.isQuickMeal ? 1 : 0,
        isCustom: 1,
        isArchived: 0,
        createdAt: now,
        updatedAt: now,
      });
    emitTableChanges('nutrition');
    return id;
  },

  /** Edits a food definition. Past meal entries keep their snapshot (never rewritten). */
  async updateFood(id: string, input: FoodInput): Promise<void> {
    await getDb()
      .update(foods)
      .set({
        name: input.name.trim() || 'Food',
        servingAmount: input.servingAmount,
        servingUnit: input.servingUnit,
        kcal: input.kcal,
        proteinG: input.proteinG,
        carbG: input.carbG,
        fatG: input.fatG,
        isQuickMeal: input.isQuickMeal ? 1 : 0,
        updatedAt: Date.now(),
      })
      .where(eq(foods.id, id));
    emitTableChanges('nutrition');
  },

  async archiveFood(id: string): Promise<void> {
    await getDb()
      .update(foods)
      .set({ isArchived: 1, updatedAt: Date.now() })
      .where(eq(foods.id, id));
    emitTableChanges('nutrition');
  },

  /** Hard-deletes a food; its meal entries survive with `food_id` SET NULL (§3.5). */
  async deleteFood(id: string): Promise<void> {
    await getDb().delete(foods).where(eq(foods.id, id));
    emitTableChanges('nutrition');
  },

  /** All meal entries for a date, in log order (DATABASE §3.5). */
  async listMealEntries(date: string): Promise<MealEntry[]> {
    const rows = await getDb()
      .select()
      .from(mealEntries)
      .where(eq(mealEntries.date, date))
      .orderBy(asc(mealEntries.loggedAt));
    return rows.map(rowToEntry);
  },

  /** Logs a meal entry with its macros already snapshotted (scaled per §4.2). */
  async addMealEntry(input: NewMealEntryInput): Promise<string> {
    const id = newId('meal');
    await getDb().insert(mealEntries).values({
      id,
      date: input.date,
      slot: input.slot,
      foodId: input.foodId,
      foodName: input.foodName,
      loggedAmount: input.loggedAmount,
      loggedUnit: input.loggedUnit,
      kcal: input.kcal,
      proteinG: input.proteinG,
      carbG: input.carbG,
      fatG: input.fatG,
      loggedAt: Date.now(),
    });
    emitTableChanges('nutrition');
    return id;
  },

  async deleteMealEntry(id: string): Promise<void> {
    await getDb().delete(mealEntries).where(eq(mealEntries.id, id));
    emitTableChanges('nutrition');
  },

  /** Re-inserts a deleted entry verbatim (Undo, UI_UX §6) — same id, same snapshot. */
  async restoreMealEntry(entry: MealEntry): Promise<void> {
    await getDb().insert(mealEntries).values({
      id: entry.id,
      date: entry.date,
      slot: entry.slot,
      foodId: entry.foodId,
      foodName: entry.foodName,
      loggedAmount: entry.loggedAmount,
      loggedUnit: entry.loggedUnit,
      kcal: entry.kcal,
      proteinG: entry.proteinG,
      carbG: entry.carbG,
      fatG: entry.fatG,
      loggedAt: entry.loggedAt,
    });
    emitTableChanges('nutrition');
  },

  /** True when any meal entry references this food (drives the delete-confirm tier). */
  async foodHasHistory(id: string): Promise<boolean> {
    const rows = await getDb()
      .select({ id: mealEntries.id })
      .from(mealEntries)
      .where(eq(mealEntries.foodId, id))
      .limit(1);
    return rows.length > 0;
  },
} as const;
