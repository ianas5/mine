import { asc, desc, eq, lte } from 'drizzle-orm';

import { emitTableChanges, getDb } from '@/core/db';
import type { MealSlot, ServingUnit } from '@/domain/nutrition';
import type { Food, MealEntry, NutritionTarget } from '@/domain/models';

import { newId } from '../id';
import { foods, mealEntries, nutritionTargets, waterDays } from '../schema/tables';

export interface TargetInput {
  readonly kcal: number;
  readonly proteinG: number;
  readonly carbG: number;
  readonly fatG: number;
  readonly waterMl: number | null;
}

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

  // ── Targets ────────────────────────────────────────────────────────────────
  // THE single, canonical target-resolution path (DATABASE §3.5). Every consumer —
  // day view, remaining, adherence, and future analytics — resolves targets here
  // and nowhere else. Do not re-implement "greatest effective_from ≤ date".

  /** The active target for a date: the row with the greatest `effective_from ≤ date`, or null. */
  async resolveTargetForDate(date: string): Promise<NutritionTarget | null> {
    const rows = await getDb()
      .select()
      .from(nutritionTargets)
      .where(lte(nutritionTargets.effectiveFrom, date))
      .orderBy(desc(nutritionTargets.effectiveFrom))
      .limit(1);
    const row = rows[0];
    return row
      ? {
          id: row.id,
          effectiveFrom: row.effectiveFrom,
          kcal: row.kcal,
          proteinG: row.proteinG,
          carbG: row.carbG,
          fatG: row.fatG,
          waterMl: row.waterMl,
        }
      : null;
  },

  /** All target eras, newest first (the read-only history in the editor, UI_UX §4.7). */
  async listTargets(): Promise<NutritionTarget[]> {
    const rows = await getDb()
      .select()
      .from(nutritionTargets)
      .orderBy(desc(nutritionTargets.effectiveFrom));
    return rows.map((row) => ({
      id: row.id,
      effectiveFrom: row.effectiveFrom,
      kcal: row.kcal,
      proteinG: row.proteinG,
      carbG: row.carbG,
      fatG: row.fatG,
      waterMl: row.waterMl,
    }));
  },

  /** "Set new targets from <date>" — writes/updates the era row for that date (§4.7). */
  async setTarget(effectiveFrom: string, input: TargetInput): Promise<void> {
    const now = Date.now();
    await getDb()
      .insert(nutritionTargets)
      .values({
        id: newId('nt'),
        effectiveFrom,
        kcal: input.kcal,
        proteinG: input.proteinG,
        carbG: input.carbG,
        fatG: input.fatG,
        waterMl: input.waterMl,
        createdAt: now,
        updatedAt: now,
      })
      .onConflictDoUpdate({
        target: nutritionTargets.effectiveFrom,
        set: {
          kcal: input.kcal,
          proteinG: input.proteinG,
          carbG: input.carbG,
          fatG: input.fatG,
          waterMl: input.waterMl,
          updatedAt: now,
        },
      });
    emitTableChanges('nutrition');
  },

  async deleteTarget(id: string): Promise<void> {
    await getDb().delete(nutritionTargets).where(eq(nutritionTargets.id, id));
    emitTableChanges('nutrition');
  },

  // ── Water ──────────────────────────────────────────────────────────────────

  /** Logged water for a date in ml, or null when the day is unlogged (0 ≠ absent, §2.4). */
  async getWater(date: string): Promise<number | null> {
    const rows = await getDb().select().from(waterDays).where(eq(waterDays.date, date));
    return rows[0]?.ml ?? null;
  },

  /** Adds (or removes, when negative) water for a date, floored at 0; logs the day. */
  async addWater(date: string, deltaMl: number): Promise<void> {
    const current = (await this.getWater(date)) ?? 0;
    const next = Math.max(0, current + deltaMl);
    await this.setWater(date, next);
  },

  /** Sets the day's water to an exact value (a logged 0 is a real value, not absence). */
  async setWater(date: string, ml: number): Promise<void> {
    const now = Date.now();
    await getDb()
      .insert(waterDays)
      .values({ date, ml: Math.max(0, ml), updatedAt: now })
      .onConflictDoUpdate({ target: waterDays.date, set: { ml: Math.max(0, ml), updatedAt: now } });
    emitTableChanges('nutrition');
  },
} as const;
