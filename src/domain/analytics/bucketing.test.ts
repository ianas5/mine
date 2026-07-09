import { addDaysIso, type IsoDate } from '@/core/utils';

import { bucketSeries, MAX_CHART_POINTS } from './bucketing';
import type { SeriesPoint } from './timeSeries';

const daily = (start: IsoDate, n: number, value = 1): SeriesPoint[] =>
  Array.from({ length: n }, (_, i) => ({ date: addDaysIso(start, i), value }));

describe('bucketSeries (ANALYTICS §4 downsampling)', () => {
  it('keeps daily granularity for 7d/30d (one point per day)', () => {
    const points = daily('2026-01-01', 30, 5);
    const out = bucketSeries(points, '30d', 'mean');
    expect(out).toHaveLength(30);
    expect(out[0]).toEqual({ date: '2026-01-01', value: 5 });
  });

  it('buckets 90d into ISO-Monday weeks', () => {
    const points = daily('2026-01-01', 90, 2);
    const out = bucketSeries(points, '90d', 'mean');
    // ~13–14 weekly buckets, each keyed to a Monday, mean of a constant series = 2.
    expect(out.length).toBeLessThanOrEqual(14);
    expect(out.every((p) => p.value === 2)).toBe(true);
  });

  it('sums within buckets for volume-style series', () => {
    const points = daily('2026-01-01', 14, 100); // 14 days × 100
    const out = bucketSeries(points, '90d', 'sum');
    const total = out.reduce((s, p) => s + p.value, 0);
    expect(total).toBe(1400); // sums are conserved across bucketing
  });

  it('coarsens to stay within the ~120-point cap for long spans', () => {
    const points = daily('2020-01-01', 2000, 1); // ~285 weeks → must coarsen to monthly
    const out = bucketSeries(points, 'all', 'mean');
    expect(out.length).toBeLessThanOrEqual(MAX_CHART_POINTS);
  });

  it('averages within a weekly bucket (mean, not sum)', () => {
    const points: SeriesPoint[] = [
      { date: '2026-01-05', value: 10 }, // Monday
      { date: '2026-01-06', value: 20 }, // same ISO week
    ];
    const out = bucketSeries(points, '90d', 'mean');
    expect(out).toEqual([{ date: '2026-01-05', value: 15 }]);
  });
});
