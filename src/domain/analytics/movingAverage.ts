import { addDaysIso } from '@/core/utils';

import { sortByDate, type SeriesPoint } from './timeSeries';

/** Points needed inside a moving-average window before it is reported (§6.2). */
export const MA_MIN_POINTS_IN_WINDOW = 2;

const mean = (points: readonly SeriesPoint[]): number =>
  points.reduce((s, p) => s + p.value, 0) / points.length;

/**
 * 7-day moving average (FITNESS_DOMAIN §6.2): for each dated point, the mean of all
 * points in the trailing 7 calendar days (inclusive). A day with fewer than 2 points
 * in its window yields no MA value (excluded, never zero-filled). De-noises daily
 * weight into a "trend weight."
 */
export function sevenDayMovingAverage(points: readonly SeriesPoint[]): SeriesPoint[] {
  const sorted = sortByDate(points);
  const out: SeriesPoint[] = [];
  for (const point of sorted) {
    const windowStart = addDaysIso(point.date, -6);
    const inWindow = sorted.filter((q) => q.date >= windowStart && q.date <= point.date);
    if (inWindow.length >= MA_MIN_POINTS_IN_WINDOW) {
      out.push({ date: point.date, value: mean(inWindow) });
    }
  }
  return out;
}

/** The latest available 7-day MA value — the headline "trend weight" — or `null`. */
export function latestMovingAverage(points: readonly SeriesPoint[]): number | null {
  const ma = sevenDayMovingAverage(points);
  return ma.length > 0 ? ma[ma.length - 1]!.value : null;
}
