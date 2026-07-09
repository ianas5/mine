import { useEffect, useMemo, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { todayIso } from '@/core/utils';
import { bodyRepository } from '@/data/repositories/bodyRepository';
import { settingsRepository } from '@/data/repositories/settingsRepository';
import {
  bucketSeries,
  computeBodyAnalytics,
  pointsInRange,
  rangeWindow,
  type BodyAnalytics,
  type RangeKey,
  type SeriesPoint,
} from '@/domain/analytics';
import type { BodySnapshot } from '@/domain/body';

interface Source {
  readonly snapshots: readonly BodySnapshot[];
  readonly targetWeightKg: number | null;
}

export interface BodyAnalyticsView {
  readonly analytics: BodyAnalytics;
  /** Bucketed in-range weight points for the chart (engine-bucketed, §4). */
  readonly weightSeries: readonly SeriesPoint[];
}

const weightPoints = (snapshots: readonly BodySnapshot[]): SeriesPoint[] =>
  snapshots.flatMap((s) => (s.weightKg !== null ? [{ date: s.date, value: s.weightKg }] : []));

/**
 * Live Body analytics for a range (ANALYTICS §5.3). Repositories provide the domain
 * models; the pure engine computes; this hook only fetches + memoizes. Recomputation is
 * keyed on the body/settings data version and the range — unchanged data recomputes
 * nothing (§7). `undefined` while first loading.
 */
export function useBodyAnalytics(range: RangeKey): BodyAnalyticsView | undefined {
  const version = useTableVersion('body', 'settings');
  const [source, setSource] = useState<Source | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void Promise.all([bodyRepository.listSnapshots(), settingsRepository.get()]).then(
      ([snapshots, settings]) => {
        if (live) setSource({ snapshots, targetWeightKg: settings.targetWeightKg });
      },
    );
    return () => {
      live = false;
    };
  }, [version]);

  return useMemo<BodyAnalyticsView | undefined>(() => {
    if (source === undefined) return undefined;

    // `listSnapshots` is newest-first, so the last element is the earliest record.
    const earliest =
      source.snapshots.length > 0 ? source.snapshots[source.snapshots.length - 1]!.date : null;
    const window = rangeWindow(range, todayIso(), earliest);

    const analytics = computeBodyAnalytics({
      snapshots: source.snapshots,
      window,
      targetWeightKg: source.targetWeightKg,
    });
    const inRangeWeight = pointsInRange(weightPoints(source.snapshots), window);
    return { analytics, weightSeries: bucketSeries(inRangeWeight, range, 'mean') };
  }, [source, range]);
}
