import { rangeWindow } from './ranges';
import { computeTrend, type TrendConfig } from './trend';
import type { SeriesPoint } from './timeSeries';

const window = rangeWindow('90d', '2026-03-01');
const cfg = (over: Partial<TrendConfig> = {}): TrendConfig => ({
  stabilityThreshold: 0.8,
  goodDirection: 'lower',
  pointNoun: 'weigh-ins',
  ...over,
});

// A clean descending series: 3 points across 28 days, −0.25/day (−7 over the span).
const descending: SeriesPoint[] = [
  { date: '2026-02-01', value: 82 },
  { date: '2026-02-15', value: 78.5 },
  { date: '2026-03-01', value: 75 },
];

describe('computeTrend (FITNESS_DOMAIN §6.4)', () => {
  it('returns no-data on an empty series', () => {
    const r = computeTrend([], cfg(), window);
    expect(r).toMatchObject({ status: 'insufficient-data', reason: 'no-data' });
  });

  it('needs at least 3 points', () => {
    const r = computeTrend(descending.slice(0, 2), cfg(), window);
    expect(r).toMatchObject({ status: 'insufficient-data', reason: 'too-few-points' });
    if (r.status === 'insufficient-data') expect(r.needed).toMatch(/1 more weigh-ins/);
  });

  it('needs a span of at least 14 days even with 3 points', () => {
    const clustered: SeriesPoint[] = [
      { date: '2026-02-25', value: 82 },
      { date: '2026-02-27', value: 81 },
      { date: '2026-03-01', value: 80 },
    ];
    const r = computeTrend(clustered, cfg(), window);
    expect(r).toMatchObject({ status: 'insufficient-data', reason: 'span-too-short' });
  });

  it('computes slope-per-week and delta over the window', () => {
    const r = computeTrend(descending, cfg(), window);
    expect(r.status).toBe('ok');
    if (r.status !== 'ok') return;
    expect(r.value.slopePerWeek).toBeCloseTo(-1.75, 2); // −0.25/day × 7
    expect(r.value.deltaOverWindow).toBeCloseTo(-7, 2); // −0.25/day × 28 days
    expect(r.computedFrom).toEqual({ points: 3, spanDays: 28 });
  });

  it('classifies a sub-threshold change as stable (deadband)', () => {
    // ~0.5 kg drop over 28 days, below the 0.8 stability threshold.
    const flat: SeriesPoint[] = [
      { date: '2026-02-01', value: 80.25 },
      { date: '2026-02-15', value: 80.0 },
      { date: '2026-03-01', value: 79.75 },
    ];
    const r = computeTrend(flat, cfg(), window);
    expect(r.status).toBe('ok');
    if (r.status !== 'ok') return;
    expect(r.value.direction).toBe('stable');
    expect(r.value.classification).toBe('stable');
  });

  it('maps decreasing → improving when lower is good (e.g. waist/weight-to-lose)', () => {
    const r = computeTrend(descending, cfg({ goodDirection: 'lower' }), window);
    if (r.status !== 'ok') throw new Error('expected ok');
    expect(r.value.direction).toBe('decreasing');
    expect(r.value.classification).toBe('improving');
  });

  it('maps decreasing → declining when higher is good (e.g. muscle/e1RM)', () => {
    const r = computeTrend(descending, cfg({ goodDirection: 'higher' }), window);
    if (r.status !== 'ok') throw new Error('expected ok');
    expect(r.value.classification).toBe('declining');
  });

  it('reports neutral classification when there is no good direction (weight, no goal)', () => {
    const r = computeTrend(descending, cfg({ goodDirection: 'neutral' }), window);
    if (r.status !== 'ok') throw new Error('expected ok');
    expect(r.value.direction).toBe('decreasing');
    expect(r.value.classification).toBe('neutral');
  });
});
