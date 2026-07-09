import { BODY_FIELDS, type BodyField, type BodySnapshot } from '@/domain/body';

import { computeBodyAnalytics } from './bodyAnalytics';
import { rangeWindow } from './ranges';

const window = rangeWindow('90d', '2026-03-01');

/** A snapshot with the given fields set and every other body field null. */
function snap(date: string, values: Partial<Record<BodyField, number>>): BodySnapshot {
  const base = Object.fromEntries(BODY_FIELDS.map((f) => [f, null])) as Record<
    BodyField,
    number | null
  >;
  return { date, ...base, ...values };
}

// Descending weight 82 → 76 over 28 days, with a dense tail so a 7-day MA exists.
const weightSnaps: BodySnapshot[] = [
  snap('2026-02-01', { weightKg: 82 }),
  snap('2026-02-15', { weightKg: 79 }),
  snap('2026-02-27', { weightKg: 76.4 }),
  snap('2026-03-01', { weightKg: 76 }),
];

describe('computeBodyAnalytics (ANALYTICS §5.3)', () => {
  it('reports the latest weight and the 7-day trend weight', () => {
    const a = computeBodyAnalytics({ snapshots: weightSnaps, window, targetWeightKg: 75 });
    expect(a.weight.latestKg).toBe(76);
    // Trailing 7 days of 2026-03-01 holds 76.4 and 76 → MA 76.2.
    expect(a.weight.trendKg).toBeCloseTo(76.2, 5);
  });

  it('classifies a falling weight as improving when above goal', () => {
    const a = computeBodyAnalytics({ snapshots: weightSnaps, window, targetWeightKg: 75 });
    expect(a.weightTrend.status).toBe('ok');
    if (a.weightTrend.status !== 'ok') return;
    expect(a.weightTrend.value.direction).toBe('decreasing');
    expect(a.weightTrend.value.classification).toBe('improving');
  });

  it('gives distance-to-target with an ETA only while heading toward goal', () => {
    const a = computeBodyAnalytics({ snapshots: weightSnaps, window, targetWeightKg: 75 });
    expect(a.distanceToTarget.status).toBe('ok');
    if (a.distanceToTarget.status !== 'ok') return;
    expect(a.distanceToTarget.value.toGoKg).toBeCloseTo(1.2, 5); // trend 76.2 − 75
    expect(a.distanceToTarget.value.ratePerWeekKg).toBeLessThan(0);
    expect(a.distanceToTarget.value.etaWeeks).not.toBeNull();
    expect(a.distanceToTarget.value.etaWeeks!).toBeGreaterThan(0);
  });

  it('returns no-target-set when no goal weight exists', () => {
    const a = computeBodyAnalytics({ snapshots: weightSnaps, window, targetWeightKg: null });
    expect(a.distanceToTarget).toMatchObject({
      status: 'insufficient-data',
      reason: 'no-target-set',
    });
    // Weight still trends, but with no good/bad verdict.
    if (a.weightTrend.status === 'ok') expect(a.weightTrend.value.classification).toBe('neutral');
  });

  it('gives no ETA when the weight trend is stable', () => {
    const flat: BodySnapshot[] = [
      snap('2026-02-01', { weightKg: 80.2 }),
      snap('2026-02-15', { weightKg: 80.0 }),
      snap('2026-02-28', { weightKg: 79.9 }),
      snap('2026-03-01', { weightKg: 80.0 }),
    ];
    const a = computeBodyAnalytics({ snapshots: flat, window, targetWeightKg: 75 });
    if (a.weightTrend.status === 'ok') expect(a.weightTrend.value.direction).toBe('stable');
    if (a.distanceToTarget.status === 'ok') expect(a.distanceToTarget.value.etaWeeks).toBeNull();
  });

  it('gives no ETA when moving away from goal', () => {
    // Gaining weight while above goal → moving away, no ETA.
    const gaining: BodySnapshot[] = [
      snap('2026-02-01', { weightKg: 78 }),
      snap('2026-02-15', { weightKg: 80 }),
      snap('2026-03-01', { weightKg: 82 }),
    ];
    const a = computeBodyAnalytics({ snapshots: gaining, window, targetWeightKg: 75 });
    if (a.distanceToTarget.status === 'ok') {
      expect(a.distanceToTarget.value.toGoKg).toBeGreaterThan(0);
      expect(a.distanceToTarget.value.etaWeeks).toBeNull();
    }
  });

  it('trends a site (waist) with lower-is-better directionality', () => {
    const waistSnaps: BodySnapshot[] = [
      snap('2026-02-01', { waistCm: 86 }),
      snap('2026-02-15', { waistCm: 84 }),
      snap('2026-03-01', { waistCm: 82 }),
    ];
    const a = computeBodyAnalytics({
      snapshots: waistSnaps,
      window,
      targetWeightKg: null,
      siteFields: ['waistCm'],
    });
    const waist = a.siteTrends.get('waistCm');
    expect(waist?.latest).toBe(82);
    expect(waist?.trend.status).toBe('ok');
    if (waist?.trend.status === 'ok') expect(waist.trend.value.classification).toBe('improving');
  });

  it('propagates insufficient-data for a site with too few measurements', () => {
    const a = computeBodyAnalytics({
      snapshots: [snap('2026-03-01', { waistCm: 82 })],
      window,
      targetWeightKg: 75,
      siteFields: ['waistCm'],
    });
    const waist = a.siteTrends.get('waistCm');
    expect(waist?.latest).toBe(82); // latest still shown
    expect(waist?.trend.status).toBe('insufficient-data');
  });
});
