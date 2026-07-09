import { daysBetweenIso, type IsoDate } from '@/core/utils';

import type { Range } from './metricResult';

/** A single `(date, value)` observation; `null`s are excluded before a series is formed. */
export interface SeriesPoint {
  readonly date: IsoDate;
  readonly value: number;
}

/** Ascending by date (FITNESS_DOMAIN §6.1). Stable, non-mutating. */
export function sortByDate(points: readonly SeriesPoint[]): SeriesPoint[] {
  return [...points].sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
}

/** The subset of points within a window (inclusive both ends); all-time keeps everything. */
export function pointsInRange(points: readonly SeriesPoint[], window: Range): SeriesPoint[] {
  return sortByDate(
    points.filter(
      (p) => (window.startDate === null || p.date >= window.startDate) && p.date <= window.endDate,
    ),
  );
}

/** Calendar-day span between the first and last point (0 for a single point). */
export function spanDays(points: readonly SeriesPoint[]): number {
  if (points.length < 2) return 0;
  const sorted = sortByDate(points);
  return daysBetweenIso(sorted[0]!.date, sorted[sorted.length - 1]!.date);
}
