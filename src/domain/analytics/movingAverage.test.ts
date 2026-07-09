import { latestMovingAverage, sevenDayMovingAverage } from './movingAverage';
import type { SeriesPoint } from './timeSeries';

describe('sevenDayMovingAverage (FITNESS_DOMAIN §6.2)', () => {
  it('averages points within the trailing 7 calendar days', () => {
    const points: SeriesPoint[] = [
      { date: '2026-01-01', value: 80 },
      { date: '2026-01-02', value: 82 },
      { date: '2026-01-03', value: 84 },
    ];
    const ma = sevenDayMovingAverage(points);
    // Day 2: mean(80,82)=81; Day 3: mean(80,82,84)=82.
    expect(ma).toEqual([
      { date: '2026-01-02', value: 81 },
      { date: '2026-01-03', value: 82 },
    ]);
  });

  it('drops days with fewer than 2 points in the window (no zero-fill)', () => {
    const points: SeriesPoint[] = [
      { date: '2026-01-01', value: 80 }, // alone → no MA
      { date: '2026-01-20', value: 84 }, // trailing 7d holds only itself → no MA
    ];
    expect(sevenDayMovingAverage(points)).toEqual([]);
  });

  it('excludes points older than 7 days from the window', () => {
    const points: SeriesPoint[] = [
      { date: '2026-01-01', value: 100 },
      { date: '2026-01-02', value: 90 },
      { date: '2026-01-10', value: 80 }, // only itself in the trailing 7 days → dropped
    ];
    const ma = sevenDayMovingAverage(points);
    expect(ma).toEqual([{ date: '2026-01-02', value: 95 }]);
  });

  it('latestMovingAverage returns the most recent MA value or null', () => {
    expect(
      latestMovingAverage([
        { date: '2026-01-01', value: 80 },
        { date: '2026-01-02', value: 82 },
      ]),
    ).toBe(81);
    expect(latestMovingAverage([{ date: '2026-01-01', value: 80 }])).toBeNull();
  });
});
