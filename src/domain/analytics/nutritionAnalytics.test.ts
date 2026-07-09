import type { TargetMacros } from '@/domain/nutrition';

import {
  calorieSkew,
  computeNutritionAnalytics,
  proteinMissStreak,
  trailingSignals,
  type DailyNutrition,
} from './nutritionAnalytics';
import { rangeWindow } from './ranges';

const TARGET: TargetMacros = { kcal: 2500, proteinG: 180, carbG: 250, fatG: 80, waterMl: 3000 };

const day = (
  date: string,
  proteinG: number,
  kcal: number,
  waterMl: number | null,
): DailyNutrition => ({
  date,
  totals: { kcal, proteinG, carbG: 250, fatG: 80 },
  target: TARGET,
  waterMl,
  logged: true,
});

const DAYS: DailyNutrition[] = [
  day('2026-03-13', 200, 2500, 3000), // all hit
  day('2026-03-14', 150, 2200, 1000), // protein + kcal under, water miss
  day('2026-03-15', 140, 2100, 500), // protein + kcal under, water miss
];

const window = rangeWindow('30d', '2026-03-15');

describe('computeNutritionAnalytics (ANALYTICS §5.2, over logged days only)', () => {
  it('computes adherence % over logged days with a target', () => {
    const a = computeNutritionAnalytics({ days: DAYS, window, today: '2026-03-15' });
    expect(a.proteinAdherence).toMatchObject({ status: 'ok' });
    if (a.proteinAdherence.status === 'ok') {
      expect(a.proteinAdherence.value).toMatchObject({ pct: 33, hitDays: 1, loggedDays: 3 });
    }
    if (a.calorieAdherence.status === 'ok') expect(a.calorieAdherence.value.pct).toBe(33);
    if (a.waterAdherence.status === 'ok') expect(a.waterAdherence.value.hitDays).toBe(1);
  });

  it('reports logging completeness as its own honest number', () => {
    const a = computeNutritionAnalytics({ days: DAYS, window, today: '2026-03-15' });
    expect(a.loggedDays).toBe(3);
    expect(a.daysInRange).toBe(30);
    expect(a.completeness).toBeCloseTo(3 / 30, 5);
  });

  it('averages macros over logged days', () => {
    const a = computeNutritionAnalytics({ days: DAYS, window, today: '2026-03-15' });
    expect(a.avg?.proteinG).toBeCloseTo((200 + 150 + 140) / 3, 5);
  });

  it('returns no-target-set water adherence when no water goal exists', () => {
    const noWater = DAYS.map((d) => ({ ...d, target: { ...TARGET, waterMl: null } }));
    const a = computeNutritionAnalytics({ days: noWater, window, today: '2026-03-15' });
    expect(a.waterAdherence).toMatchObject({
      status: 'insufficient-data',
      reason: 'no-target-set',
    });
  });
});

describe('insight signal helpers', () => {
  it('counts consecutive logged protein-miss days from the newest', () => {
    expect(proteinMissStreak(DAYS)).toBe(2); // 03-15 miss, 03-14 miss, 03-13 hit
  });

  it('skews calorie misses to one side when ≥ 70% fall there', () => {
    expect(calorieSkew(DAYS, '2026-03-15', 14)).toBe('under'); // both misses are under
  });

  it('summarizes the trailing 7 days for insight rules', () => {
    const s = trailingSignals(DAYS, '2026-03-15', 7);
    expect(s.loggedCount).toBe(3);
    expect(s.proteinHits).toBe(1);
    expect(s.waterHits).toBe(1);
    expect(s.waterTargetDays).toBe(3);
  });
});
