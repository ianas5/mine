import { pointsInRange } from './timeSeries';
import { rangeWindow } from './ranges';
import type { SeriesPoint } from './timeSeries';

describe('rangeWindow (FITNESS_DOMAIN §7)', () => {
  it('makes a 7-day window inclusive of today and the six days before', () => {
    const w = rangeWindow('7d', '2026-03-10');
    expect(w).toEqual({ key: '7d', startDate: '2026-03-04', endDate: '2026-03-10', days: 7 });
  });

  it('opens all-time from the first record (or null when empty)', () => {
    expect(rangeWindow('all', '2026-03-10', '2025-01-01')).toEqual({
      key: 'all',
      startDate: '2025-01-01',
      endDate: '2026-03-10',
      days: null,
    });
    expect(rangeWindow('all', '2026-03-10', null).startDate).toBeNull();
  });
});

describe('pointsInRange', () => {
  const points: SeriesPoint[] = [
    { date: '2026-02-01', value: 1 },
    { date: '2026-03-04', value: 2 },
    { date: '2026-03-10', value: 3 },
  ];

  it('keeps only points within the window, sorted ascending', () => {
    const w = rangeWindow('7d', '2026-03-10');
    expect(pointsInRange(points, w).map((p) => p.date)).toEqual(['2026-03-04', '2026-03-10']);
  });

  it('keeps everything for an all-time window with no start bound', () => {
    const w = rangeWindow('all', '2026-03-10', null);
    expect(pointsInRange(points, w)).toHaveLength(3);
  });
});
