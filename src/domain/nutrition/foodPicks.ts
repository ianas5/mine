import type { MealSlot } from './taxonomy';

/**
 * The Log Meal picker's smart-default heuristic (UI_UX §5.2), a pure function over the
 * user's own history — no AI, just transparent frequency. Retrieval stays in the
 * repository; every ranking/last-used/most-frequent decision lives here (ARCHITECTURE
 * §9.1, ANALYTICS rule 9). Every default it produces is a pre-fill the UI shows and the
 * user can override in one tap — it never blocks, asks, or locks a choice in.
 */

/** One recent meal-log row the heuristic reads (the repository supplies these). */
export interface MealUsageRow {
  readonly foodId: string | null;
  readonly loggedAmount: number;
  readonly slot: MealSlot | null;
  readonly loggedAt: number;
}

/** Per-food usage summary: how often, the last-used portion, and the habitual slot. */
export interface FoodUsage {
  readonly count: number;
  readonly lastAmount: number | null;
  readonly lastLoggedAt: number;
  /** The habitual slot for this food, or `null` on a tie or no history — the caller
   * then falls back to the time-of-day default so a tie never surprises the user. */
  readonly slot: MealSlot | null;
}

/**
 * The slot a food is *most often* logged in (§5.2). Returns `null` when there is no
 * history **or the lead is tied** — a tie is not a habit, so the caller defers to the
 * time-of-day fallback rather than guessing one arbitrarily.
 */
export function mostFrequentSlot(slots: readonly (MealSlot | null)[]): MealSlot | null {
  const counts = new Map<MealSlot, number>();
  for (const s of slots) if (s !== null) counts.set(s, (counts.get(s) ?? 0) + 1);
  let maxCount = 0;
  for (const c of counts.values()) if (c > maxCount) maxCount = c;
  if (maxCount === 0) return null;
  const leaders = [...counts.entries()].filter(([, c]) => c === maxCount);
  return leaders.length === 1 ? leaders[0]![0] : null;
}

/**
 * Folds recent meal-log rows into per-food usage. Order-independent: the last-used
 * portion is the one from the row with the greatest `loggedAt`, not "whichever came
 * first," so the result never depends on the query's row order.
 */
export function aggregateFoodUsage(rows: readonly MealUsageRow[]): Map<string, FoodUsage> {
  interface Acc {
    count: number;
    lastAmount: number | null;
    lastLoggedAt: number;
    slots: (MealSlot | null)[];
  }
  const acc = new Map<string, Acc>();
  for (const r of rows) {
    if (r.foodId === null) continue;
    const u = acc.get(r.foodId);
    if (u) {
      u.count += 1;
      u.slots.push(r.slot);
      if (r.loggedAt >= u.lastLoggedAt) {
        u.lastLoggedAt = r.loggedAt;
        u.lastAmount = r.loggedAmount;
      }
    } else {
      acc.set(r.foodId, {
        count: 1,
        lastAmount: r.loggedAmount,
        lastLoggedAt: r.loggedAt,
        slots: [r.slot],
      });
    }
  }
  const out = new Map<string, FoodUsage>();
  for (const [id, u] of acc) {
    out.set(id, {
      count: u.count,
      lastAmount: u.lastAmount,
      lastLoggedAt: u.lastLoggedAt,
      slot: mostFrequentSlot(u.slots),
    });
  }
  return out;
}

/** The minimum a food needs to be ranked and pre-filled. */
export interface RankableFood {
  readonly id: string;
  readonly name: string;
  readonly isQuickMeal: boolean;
}

export interface FoodPickOf<F> {
  readonly food: F;
  readonly lastAmount: number | null;
  readonly slot: MealSlot | null;
  readonly useCount: number;
}

/**
 * Orders foods for the picker (§4.3/§5.2): quick meals pinned, then most-used, then
 * most-recently-used, then alphabetical. With no history every food falls to the
 * alphabetical tail — the sensible static default (`history → static default → ask
 * nothing`). Pure and stable; carries each food's last-used portion + habitual slot.
 */
export function orderFoodPicks<F extends RankableFood>(
  foods: readonly F[],
  usage: ReadonlyMap<string, FoodUsage>,
): FoodPickOf<F>[] {
  return foods
    .map((food) => {
      const u = usage.get(food.id);
      return {
        food,
        lastAmount: u?.lastAmount ?? null,
        slot: u?.slot ?? null,
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
}
