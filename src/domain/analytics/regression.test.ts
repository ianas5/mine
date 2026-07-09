import { linearRegression } from './regression';
import type { SeriesPoint } from './timeSeries';

describe('linearRegression (FITNESS_DOMAIN §6.4 least-squares)', () => {
  it('recovers a known slope in units per day', () => {
    // +1.0 per day starting at 100 on the base date.
    const points: SeriesPoint[] = [
      { date: '2026-01-01', value: 100 },
      { date: '2026-01-02', value: 101 },
      { date: '2026-01-03', value: 102 },
      { date: '2026-01-04', value: 103 },
    ];
    const r = linearRegression(points)!;
    expect(r.slopePerDay).toBeCloseTo(1, 6);
    expect(r.intercept).toBeCloseTo(100, 6);
  });

  it('respects real day gaps (not evenly-spaced samples)', () => {
    // 100 on day 0, 110 on day 10 → slope 1.0/day, regardless of only two readings.
    const r = linearRegression([
      { date: '2026-01-01', value: 100 },
      { date: '2026-01-11', value: 110 },
    ])!;
    expect(r.slopePerDay).toBeCloseTo(1, 6);
  });

  it('returns null for fewer than two points', () => {
    expect(linearRegression([{ date: '2026-01-01', value: 100 }])).toBeNull();
    expect(linearRegression([])).toBeNull();
  });

  it('returns null when every point shares one date (no x-variance)', () => {
    expect(
      linearRegression([
        { date: '2026-01-01', value: 100 },
        { date: '2026-01-01', value: 104 },
      ]),
    ).toBeNull();
  });
});
