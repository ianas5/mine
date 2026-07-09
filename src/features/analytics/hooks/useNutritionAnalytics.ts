import { useEffect, useMemo, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { addDaysIso, todayIso } from '@/core/utils';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';
import {
  computeNutritionAnalytics,
  rangeWindow,
  type DailyNutrition,
  type NutritionAnalytics,
  type RangeKey,
} from '@/domain/analytics';

/**
 * Nutrition analytics for a range (ANALYTICS §5.2). Repositories fetch per-day rows (with
 * targets resolved through the canonical path); the pure calculator computes over logged
 * days only. Memoized by data version + range. `undefined` while first loading.
 */
export function useNutritionAnalytics(range: RangeKey): NutritionAnalytics | undefined {
  const version = useTableVersion('nutrition');
  const [days, setDays] = useState<readonly DailyNutrition[] | undefined>(undefined);

  useEffect(() => {
    let live = true;
    const since = range === 'all' ? '2000-01-01' : addDaysIso(todayIso(), -400);
    void nutritionRepository.getDailyNutritionSince(since).then((rows) => {
      if (live) setDays(rows);
    });
    return () => {
      live = false;
    };
  }, [version, range]);

  return useMemo<NutritionAnalytics | undefined>(() => {
    if (days === undefined) return undefined;
    const today = todayIso();
    return computeNutritionAnalytics({ days, window: rangeWindow(range, today), today });
  }, [days, range]);
}
