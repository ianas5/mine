import { addDaysIso, isoWeekday, type IsoDate } from '@/core/utils';

import type { RangeKey } from './metricResult';
import { sortByDate, type SeriesPoint } from './timeSeries';

/** Chart series never exceed ~120 points (ANALYTICS §4); bucketing downsamples for display. */
export const MAX_CHART_POINTS = 120;

export type Granularity = 'daily' | 'weekly' | 'monthly';
/** Means for measurements; sums for counts/volume (§4). */
export type Aggregation = 'mean' | 'sum';

const COARSER: Record<Granularity, Granularity | null> = {
  daily: 'weekly',
  weekly: 'monthly',
  monthly: null,
};

function baseGranularity(range: RangeKey): Granularity {
  if (range === '7d' || range === '30d') return 'daily';
  return 'weekly'; // 90d/180d/365d/all start weekly, then coarsen if still > 120 pts
}

/** The bucket a date falls in: the day, its ISO-Monday week start, or its month start. */
function bucketKey(date: IsoDate, granularity: Granularity): IsoDate {
  if (granularity === 'daily') return date;
  if (granularity === 'weekly') return addDaysIso(date, -isoWeekday(date)); // back to Monday
  return `${date.slice(0, 7)}-01` as IsoDate; // month start
}

function aggregateAt(
  points: readonly SeriesPoint[],
  granularity: Granularity,
  aggregation: Aggregation,
): SeriesPoint[] {
  const groups = new Map<IsoDate, number[]>();
  for (const point of points) {
    const key = bucketKey(point.date, granularity);
    const bucket = groups.get(key);
    if (bucket) bucket.push(point.value);
    else groups.set(key, [point.value]);
  }
  return [...groups.entries()]
    .map(([date, values]) => {
      const total = values.reduce((s, v) => s + v, 0);
      return { date, value: aggregation === 'sum' ? total : total / values.length };
    })
    .sort((a, b) => (a.date < b.date ? -1 : 1));
}

/**
 * Downsamples a raw series into ≤ 120 display points at the granularity for its range
 * (daily 7d/30d; weekly 90d/180d/365d; coarsened to monthly for very long all-time
 * spans). **Display only** — trend math always runs on the raw points (§4).
 */
export function bucketSeries(
  points: readonly SeriesPoint[],
  range: RangeKey,
  aggregation: Aggregation,
): SeriesPoint[] {
  const sorted = sortByDate(points);
  let granularity = baseGranularity(range);
  let result = aggregateAt(sorted, granularity, aggregation);
  let coarser = COARSER[granularity];
  while (result.length > MAX_CHART_POINTS && coarser !== null) {
    granularity = coarser;
    result = aggregateAt(sorted, granularity, aggregation);
    coarser = COARSER[granularity];
  }
  return result;
}
