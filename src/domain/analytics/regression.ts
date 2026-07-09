import { daysBetweenIso } from '@/core/utils';

import type { SeriesPoint } from './timeSeries';

export interface Regression {
  /** Least-squares slope in value-units per day. */
  readonly slopePerDay: number;
  /** Value at the first point's day (day-index 0). */
  readonly intercept: number;
}

/**
 * Least-squares linear regression of value vs. **day-index** (FITNESS_DOMAIN §6.4).
 * Day-index is calendar days since the earliest point, so real gaps between weigh-ins
 * are respected (not treated as evenly spaced samples). Returns `null` when there are
 * fewer than two points or every point shares one date (zero x-variance → no slope).
 */
export function linearRegression(points: readonly SeriesPoint[]): Regression | null {
  if (points.length < 2) return null;

  const base = points.reduce((min, p) => (p.date < min ? p.date : min), points[0]!.date);
  const xs = points.map((p) => daysBetweenIso(base, p.date));
  const ys = points.map((p) => p.value);
  const n = points.length;

  const meanX = xs.reduce((s, x) => s + x, 0) / n;
  const meanY = ys.reduce((s, y) => s + y, 0) / n;

  let num = 0;
  let den = 0;
  for (let i = 0; i < n; i += 1) {
    const dx = xs[i]! - meanX;
    num += dx * (ys[i]! - meanY);
    den += dx * dx;
  }
  if (den === 0) return null; // all points on the same day → undefined slope

  const slopePerDay = num / den;
  return { slopePerDay, intercept: meanY - slopePerDay * meanX };
}
