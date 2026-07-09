import { addDaysIso, type IsoDate } from '@/core/utils';

import type { Range, RangeKey } from './metricResult';

/** The six canonical rolling ranges ending today (FITNESS_DOMAIN §7, ANALYTICS §4). */
export const RANGE_KEYS: readonly RangeKey[] = ['7d', '30d', '90d', '180d', '365d', 'all'];

export const RANGE_LABELS: Record<RangeKey, string> = {
  '7d': '7D',
  '30d': '30D',
  '90d': '3M',
  '180d': '6M',
  '365d': '1Y',
  all: 'All',
};

/** Fixed window lengths in days; all-time is open-ended (first record → today). */
const RANGE_DAYS: Record<Exclude<RangeKey, 'all'>, number> = {
  '7d': 7,
  '30d': 30,
  '90d': 90,
  '180d': 180,
  '365d': 365,
};

/**
 * Resolves a range key to a concrete window ending `today`. For all-time the start is
 * the earliest record (or `null` when there is no data); the window is inclusive of
 * both ends. Pure — `today` and `firstRecord` are inputs (ANALYTICS rule 8).
 */
export function rangeWindow(
  key: RangeKey,
  today: IsoDate,
  firstRecord: IsoDate | null = null,
): Range {
  if (key === 'all') {
    return { key, startDate: firstRecord, endDate: today, days: null };
  }
  const days = RANGE_DAYS[key];
  // Inclusive of today, so a 7-day window spans today and the six days before it.
  return { key, startDate: addDaysIso(today, -(days - 1)), endDate: today, days };
}
