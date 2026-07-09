import { addDaysIso } from '@/core/utils';

import { evaluateInsights, selectDashboardInsights } from './engine';
import {
  baseContext,
  distance,
  loggedDay,
  recompFired,
  siteMap,
  TODAY,
  trend,
} from './insightTestKit';
import type { CooldownMap, InsightContext } from './types';

const ids = (ctx: InsightContext, cooldowns?: CooldownMap): string[] =>
  evaluateInsights(ctx, cooldowns).map((i) => i.ruleId);
const fires = (ctx: InsightContext, ruleId: string): boolean => ids(ctx).includes(ruleId);

// Nutrition-day generators over the trailing window ending today.
const days = (
  n: number,
  make: (i: number) => Partial<import('../nutritionAnalytics').DailyNutrition>,
) => Array.from({ length: n }, (_, i) => loggedDay({ date: addDaysIso(TODAY, -i), ...make(i) }));

describe('insight engine — quiet state', () => {
  it('fires nothing from an empty/quiet context', () => {
    expect(evaluateInsights(baseContext())).toEqual([]);
  });
});

describe('insight rules — each triggers on its condition (§6.2)', () => {
  it('1 recomp-signal', () => {
    expect(fires(baseContext({ recomp: recompFired() }), 'recomp-signal')).toBe(true);
  });

  it('2 weight-trend-to-target', () => {
    const ctx = baseContext({
      body: {
        weight: { latestKg: 80, trendKg: 80 },
        weightTrend: trend('improving', 'decreasing', -0.4),
        distanceToTarget: distance({ atGoal: false }),
        siteTrends: new Map(),
      },
    });
    expect(fires(ctx, 'weight-trend-to-target')).toBe(true);
    expect(fires(ctx, 'weight-trend-away')).toBe(false); // conflict guard (2 vs 3)
  });

  it('3 weight-trend-away', () => {
    const ctx = baseContext({
      body: {
        weight: { latestKg: 80, trendKg: 80 },
        weightTrend: trend('declining', 'increasing', 0.4),
        distanceToTarget: distance({}),
        siteTrends: new Map(),
      },
    });
    expect(fires(ctx, 'weight-trend-away')).toBe(true);
    expect(fires(ctx, 'weight-trend-to-target')).toBe(false);
  });

  it('4 weight-stalled', () => {
    const ctx = baseContext({
      body: {
        weight: { latestKg: 80, trendKg: 80 },
        weightTrend: trend('stable', 'stable', 0, 30),
        distanceToTarget: distance({ toGoKg: 5, atGoal: false }),
        siteTrends: new Map(),
      },
      nutrition: {
        ...baseContext().nutrition,
        calorieAdherence: {
          status: 'ok',
          value: { pct: 40, hitDays: 4, loggedDays: 10 },
          window: baseContext().window,
          computedFrom: { points: 10, spanDays: 0 },
        },
      },
    });
    expect(fires(ctx, 'weight-stalled')).toBe(true);
  });

  it('5 waist-decreasing', () => {
    const ctx = baseContext({
      body: {
        ...baseContext().body,
        siteTrends: siteMap([['waistCm', trend('improving', 'decreasing', -0.3)]]),
      },
    });
    expect(fires(ctx, 'waist-decreasing')).toBe(true);
  });

  it('6 site-growing and 7 site-declining are exclusive per site', () => {
    const grow = baseContext({
      body: {
        ...baseContext().body,
        siteTrends: siteMap([['chestCm', trend('improving', 'increasing', 0.2)]]),
      },
    });
    expect(fires(grow, 'site-growing')).toBe(true);
    expect(fires(grow, 'site-declining')).toBe(false);

    const decline = baseContext({
      body: {
        ...baseContext().body,
        siteTrends: siteMap([['chestCm', trend('declining', 'decreasing', -0.2)]]),
      },
    });
    expect(fires(decline, 'site-declining')).toBe(true);
    expect(fires(decline, 'site-growing')).toBe(false);
  });

  it('8 protein-miss-streak', () => {
    const ctx = baseContext({
      nutritionDays: days(3, () => ({
        totals: { kcal: 2000, proteinG: 100, carbG: 200, fatG: 60 },
      })),
    });
    expect(fires(ctx, 'protein-miss-streak')).toBe(true);
  });

  it('9 protein-strong-week', () => {
    const ctx = baseContext({
      nutritionDays: days(6, () => ({
        totals: { kcal: 2000, proteinG: 200, carbG: 200, fatG: 60 },
      })),
    });
    expect(fires(ctx, 'protein-strong-week')).toBe(true);
  });

  it('10 kcal-skew', () => {
    const ctx = baseContext({
      nutritionDays: days(10, () => ({
        totals: { kcal: 1500, proteinG: 200, carbG: 200, fatG: 60 },
      })),
    });
    expect(fires(ctx, 'kcal-skew')).toBe(true);
  });

  it('11 logging-gap', () => {
    const ctx = baseContext({
      nutritionDays: days(2, () => ({
        totals: { kcal: 2000, proteinG: 200, carbG: 200, fatG: 60 },
      })),
    });
    expect(fires(ctx, 'logging-gap')).toBe(true);
  });

  it('12 consistency-up / 13 consistency-down are exclusive', () => {
    const up = baseContext({ completedWeekConsistency: { current: 100, previous: 60 } });
    expect(fires(up, 'consistency-up')).toBe(true);
    expect(fires(up, 'consistency-down')).toBe(false);

    const down = baseContext({ completedWeekConsistency: { current: 50, previous: 100 } });
    expect(fires(down, 'consistency-down')).toBe(true);
    expect(fires(down, 'consistency-up')).toBe(false);
  });

  it('14 streak-milestone at 4 weeks (and not at 5)', () => {
    const w = baseContext().workout;
    const at4 = baseContext({
      workout: { ...w, consistency: { progress: { completed: 4, planned: 4 }, streak: 4 } },
    });
    expect(fires(at4, 'streak-milestone')).toBe(true);
    const at5 = baseContext({
      workout: { ...w, consistency: { progress: { completed: 4, planned: 4 }, streak: 5 } },
    });
    expect(fires(at5, 'streak-milestone')).toBe(false);
  });

  it('15 new-pr', () => {
    const ctx = baseContext({ recentPrs: [{ exerciseId: 'ex1', name: 'Bench', kinds: ['e1rm'] }] });
    expect(fires(ctx, 'new-pr')).toBe(true);
  });

  it('16 strength-trend-up / 17 strength-trend-down are exclusive per exercise', () => {
    const w = baseContext().workout;
    const up = baseContext({
      workout: {
        ...w,
        keyExercises: [
          {
            exerciseId: 'ex1',
            name: 'Bench',
            sessions: 6,
            trend: trend('improving', 'increasing', 1),
          },
        ],
      },
    });
    expect(fires(up, 'strength-trend-up')).toBe(true);
    expect(fires(up, 'strength-trend-down')).toBe(false);
  });

  it('18 push-pull-imbalance only with ≥ 8 sessions', () => {
    const w = baseContext().workout;
    const flagged = {
      ...w,
      pushPull: { ratio: 1.7, numeratorKg: 170, denominatorKg: 100, flagged: true },
    };
    expect(fires(baseContext({ workout: flagged, sessions30d: 4 }), 'push-pull-imbalance')).toBe(
      false,
    );
    expect(fires(baseContext({ workout: flagged, sessions30d: 10 }), 'push-pull-imbalance')).toBe(
      true,
    );
  });

  it('19 neglected-muscle only with ≥ 8 sessions', () => {
    expect(
      fires(baseContext({ neglectedGroups: ['back'], sessions30d: 10 }), 'neglected-muscle'),
    ).toBe(true);
    expect(
      fires(baseContext({ neglectedGroups: ['back'], sessions30d: 4 }), 'neglected-muscle'),
    ).toBe(false);
  });

  it('20 volume-drop', () => {
    const w = baseContext().workout;
    const series = [
      { date: '2026-02-09', value: 10000 },
      { date: '2026-02-16', value: 10000 },
      { date: '2026-02-23', value: 10000 },
      { date: '2026-03-02', value: 10000 },
      { date: '2026-03-09', value: 3000 }, // 30% of the 10000 average
    ];
    expect(fires(baseContext({ workout: { ...w, volumeSeries: series } }), 'volume-drop')).toBe(
      true,
    );
  });

  it('21 measurement-due / 22 photo-due when overdue', () => {
    const ctx = baseContext({
      lastSnapshotDate: addDaysIso(TODAY, -30),
      lastPhotoDate: addDaysIso(TODAY, -40),
    });
    expect(fires(ctx, 'measurement-due')).toBe(true);
    expect(fires(ctx, 'photo-due')).toBe(true);
  });

  it('23 water-low-week', () => {
    const ctx = baseContext({ nutritionDays: days(5, () => ({ waterMl: 0 })) });
    expect(fires(ctx, 'water-low-week')).toBe(true);
  });
});

describe('cooldown (§6.3)', () => {
  it('suppresses a fired instance within its cooldown window', () => {
    const ctx = baseContext({ recomp: recompFired() });
    const cooldowns: CooldownMap = {
      'recomp-signal': { lastFired: addDaysIso(TODAY, -3), classification: 'recomp' },
    };
    expect(ids(ctx, cooldowns)).not.toContain('recomp-signal');
  });

  it('breaks cooldown when the classification flips', () => {
    const ctx = baseContext({
      body: {
        ...baseContext().body,
        siteTrends: siteMap([['waistCm', trend('improving', 'decreasing', -0.3)]]),
      },
    });
    const cooldowns: CooldownMap = {
      'waist-decreasing': { lastFired: addDaysIso(TODAY, -1), classification: 'declining' }, // stale → flip
    };
    expect(ids(ctx, cooldowns)).toContain('waist-decreasing');
  });
});

describe('dashboard selection (§6.3)', () => {
  it('caps at 3, ≤ 2 per category, ≤ 1 housekeeping', () => {
    const ctx = baseContext({
      recomp: recompFired(), // body
      body: {
        weight: { latestKg: 80, trendKg: 80 },
        weightTrend: trend('improving', 'decreasing', -0.4), // body
        distanceToTarget: distance({}),
        siteTrends: siteMap([['waistCm', trend('improving', 'decreasing', -0.3)]]), // body
      },
      lastSnapshotDate: addDaysIso(TODAY, -30), // housekeeping
      lastPhotoDate: addDaysIso(TODAY, -40), // housekeeping
      completedWeekConsistency: { current: 100, previous: 50 }, // consistency
    });
    const dash = selectDashboardInsights(evaluateInsights(ctx));
    expect(dash.length).toBeLessThanOrEqual(3);
    expect(dash.filter((i) => i.category === 'body').length).toBeLessThanOrEqual(2);
    expect(dash.filter((i) => i.category === 'housekeeping').length).toBeLessThanOrEqual(1);
  });
});
