import {
  aggregateFoodUsage,
  mostFrequentSlot,
  orderFoodPicks,
  type MealUsageRow,
  type RankableFood,
} from './foodPicks';

describe('mostFrequentSlot (§5.2 — habit, or defer on a tie)', () => {
  it('returns the strict most-frequent slot', () => {
    expect(mostFrequentSlot(['lunch', 'lunch', 'dinner'])).toBe('lunch');
  });

  it('returns null on a tie — a tie is not a habit, so defer to the time-of-day fallback', () => {
    expect(mostFrequentSlot(['lunch', 'dinner'])).toBeNull();
  });

  it('returns null with no history', () => {
    expect(mostFrequentSlot([])).toBeNull();
    expect(mostFrequentSlot([null, null])).toBeNull();
  });

  it('ignores unslotted entries', () => {
    expect(mostFrequentSlot(['breakfast', null, 'breakfast'])).toBe('breakfast');
  });
});

describe('aggregateFoodUsage', () => {
  const row = (
    foodId: string | null,
    loggedAmount: number,
    slot: string | null,
    loggedAt: number,
  ): MealUsageRow => ({ foodId, loggedAmount, slot: slot as MealUsageRow['slot'], loggedAt });

  it('counts uses and takes the last-used portion from the most recent row (order-independent)', () => {
    // Deliberately out of order: the newest row (loggedAt 300, amount 250) must win.
    const usage = aggregateFoodUsage([
      row('a', 100, 'lunch', 100),
      row('a', 250, 'dinner', 300),
      row('a', 200, 'lunch', 200),
    ]);
    const a = usage.get('a')!;
    expect(a.count).toBe(3);
    expect(a.lastAmount).toBe(250); // amount at the greatest loggedAt, not first-seen
    expect(a.slot).toBe('lunch'); // 2× lunch vs 1× dinner
  });

  it('skips rows with no foodId (a deleted food keeps history but is unrankable)', () => {
    const usage = aggregateFoodUsage([row(null, 100, 'lunch', 100), row('b', 50, 'snacks', 90)]);
    expect(usage.has('b')).toBe(true);
    expect(usage.size).toBe(1);
  });
});

describe('orderFoodPicks (§4.3/§5.2 ordering + fallback chain)', () => {
  const food = (id: string, name: string, isQuickMeal = false): RankableFood => ({
    id,
    name,
    isQuickMeal,
  });

  it('pins quick meals, then most-used, then most-recent, then alphabetical', () => {
    const foods = [
      food('staple', 'Chicken'),
      food('rare', 'Anchovies'),
      food('quick', 'Protein Shake', true),
    ];
    const usage = aggregateFoodUsage([
      { foodId: 'staple', loggedAmount: 200, slot: 'lunch', loggedAt: 500 },
      { foodId: 'staple', loggedAmount: 200, slot: 'lunch', loggedAt: 400 },
      { foodId: 'rare', loggedAmount: 10, slot: 'dinner', loggedAt: 100 },
    ]);
    const order = orderFoodPicks(foods, usage).map((p) => p.food.id);
    expect(order).toEqual(['quick', 'staple', 'rare']); // quick pinned even with 0 uses
  });

  it('with no history falls back to the alphabetical static default, slot/amount null', () => {
    const foods = [food('b', 'Banana'), food('a', 'Apple')];
    const picks = orderFoodPicks(foods, aggregateFoodUsage([]));
    expect(picks.map((p) => p.food.id)).toEqual(['a', 'b']);
    expect(picks[0]!.slot).toBeNull();
    expect(picks[0]!.lastAmount).toBeNull();
    expect(picks[0]!.useCount).toBe(0);
  });

  it('breaks equal use-counts by recency', () => {
    const foods = [food('older', 'Older'), food('newer', 'Newer')];
    const usage = aggregateFoodUsage([
      { foodId: 'older', loggedAmount: 1, slot: null, loggedAt: 100 },
      { foodId: 'newer', loggedAmount: 1, slot: null, loggedAt: 900 },
    ]);
    expect(orderFoodPicks(foods, usage).map((p) => p.food.id)).toEqual(['newer', 'older']);
  });
});
