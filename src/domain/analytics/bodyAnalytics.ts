import { BODY_DIRECTION, BODY_STABILITY, type BodyField, type BodySnapshot } from '@/domain/body';

import { insufficient, ok, type MetricResult, type Range } from './metricResult';
import { latestMovingAverage } from './movingAverage';
import { sortByDate, type SeriesPoint } from './timeSeries';
import { computeTrend, type GoodDirection, type Trend } from './trend';

/** Weight stability deadband (kg) — the §6.4 default, reused for the ETA guard. */
export const WEIGHT_STABILITY_KG = BODY_STABILITY.weightKg;

const round1 = (n: number): number => Math.round(n * 10) / 10;

export interface WeightHeadline {
  /** Most recent raw weigh-in ("latest"). */
  readonly latestKg: number | null;
  /** Latest 7-day moving average ("trend weight"), the de-noised headline (§5.3). */
  readonly trendKg: number | null;
}

export interface DistanceToTarget {
  readonly targetKg: number;
  /** Current trend weight − target: positive = above goal, negative = below. */
  readonly toGoKg: number;
  /** Current weekly rate of change (from the weight trend), or null if no trend. */
  readonly ratePerWeekKg: number | null;
  /** Honest ETA in weeks — only when the trend is meaningful AND moving toward goal. */
  readonly etaWeeks: number | null;
  /** Within the stability deadband of the goal → effectively there. */
  readonly atGoal: boolean;
}

export interface SiteAnalysis {
  /** Most recent in-range value for the site (for the headline value). */
  readonly latest: number | null;
  readonly trend: MetricResult<Trend>;
}

export interface BodyAnalytics {
  readonly weight: WeightHeadline;
  readonly weightTrend: MetricResult<Trend>;
  readonly distanceToTarget: MetricResult<DistanceToTarget>;
  readonly siteTrends: ReadonlyMap<BodyField, SiteAnalysis>;
}

export interface BodyAnalyticsInput {
  readonly snapshots: readonly BodySnapshot[];
  readonly window: Range;
  readonly targetWeightKg: number | null;
  /** Circumference/composition sites to trend (default: waist). */
  readonly siteFields?: readonly BodyField[];
}

function fieldSeries(snapshots: readonly BodySnapshot[], field: BodyField): SeriesPoint[] {
  const points: SeriesPoint[] = [];
  for (const snapshot of snapshots) {
    const value = snapshot[field];
    if (value !== null) points.push({ date: snapshot.date, value });
  }
  return points;
}

/** Which way is "good" for weight given the goal (§5.3): toward the target, else neutral. */
function weightGoodDirection(current: number | null, target: number | null): GoodDirection {
  if (target === null || current === null) return 'neutral';
  const diff = current - target;
  if (Math.abs(diff) < WEIGHT_STABILITY_KG) return 'neutral'; // essentially at goal → don't judge
  return diff > 0 ? 'lower' : 'higher';
}

function distanceToTarget(
  current: number | null,
  target: number | null,
  weightTrend: MetricResult<Trend>,
  window: Range,
): MetricResult<DistanceToTarget> {
  if (target === null) {
    return insufficient(
      'no-target-set',
      'Set a goal weight in Settings to track distance to target',
    );
  }
  if (current === null) {
    return insufficient('no-data', 'Log your weight to track distance to target');
  }

  const toGoKg = round1(current - target);
  const atGoal = Math.abs(current - target) < WEIGHT_STABILITY_KG;
  const ratePerWeekKg = weightTrend.status === 'ok' ? weightTrend.value.slopePerWeek : null;

  let etaWeeks: number | null = null;
  if (
    !atGoal &&
    weightTrend.status === 'ok' &&
    weightTrend.value.direction !== 'stable' &&
    ratePerWeekKg !== null &&
    ratePerWeekKg !== 0
  ) {
    const movingToward = (toGoKg > 0 && ratePerWeekKg < 0) || (toGoKg < 0 && ratePerWeekKg > 0);
    if (movingToward)
      etaWeeks = Math.max(1, Math.round(Math.abs(toGoKg) / Math.abs(ratePerWeekKg)));
  }

  const value: DistanceToTarget = { targetKg: target, toGoKg, ratePerWeekKg, etaWeeks, atGoal };
  return ok(value, window, { points: 1, spanDays: 0 });
}

/**
 * The Body analytics calculator (ANALYTICS §5.3), pure. Consumes windowed snapshots +
 * the goal weight and returns the weight headline (latest + 7-day trend weight), the
 * weight trend classified toward/away from goal, an honest distance-to-target (ETA only
 * when the trend is meaningful and heading the right way), and per-site trends. Every
 * result is a `MetricResult` — insufficient data is stated, never fabricated.
 */
export function computeBodyAnalytics(input: BodyAnalyticsInput): BodyAnalytics {
  const windowed = pointsAwareSnapshots(input.snapshots, input.window);
  const weightSeries = fieldSeries(windowed, 'weightKg');
  const sortedWeight = sortByDate(weightSeries);

  const latestKg = sortedWeight.length > 0 ? sortedWeight[sortedWeight.length - 1]!.value : null;
  const trendKg = latestMovingAverage(weightSeries);
  const current = trendKg ?? latestKg;

  const weightTrend = computeTrend(
    weightSeries,
    {
      stabilityThreshold: WEIGHT_STABILITY_KG,
      goodDirection: weightGoodDirection(current, input.targetWeightKg),
      pointNoun: 'weigh-ins',
    },
    input.window,
  );

  const siteTrends = new Map<BodyField, SiteAnalysis>();
  for (const field of input.siteFields ?? (['waistCm'] as const)) {
    const series = sortByDate(fieldSeries(windowed, field));
    siteTrends.set(field, {
      latest: series.length > 0 ? series[series.length - 1]!.value : null,
      trend: computeTrend(
        series,
        {
          stabilityThreshold: BODY_STABILITY[field],
          goodDirection: BODY_DIRECTION[field],
          pointNoun: 'measurements',
        },
        input.window,
      ),
    });
  }

  return {
    weight: { latestKg, trendKg },
    weightTrend,
    distanceToTarget: distanceToTarget(current, input.targetWeightKg, weightTrend, input.window),
    siteTrends,
  };
}

/** Snapshots restricted to the window (belt-and-braces; hooks pass windowed rows). */
function pointsAwareSnapshots(
  snapshots: readonly BodySnapshot[],
  window: Range,
): readonly BodySnapshot[] {
  return snapshots.filter(
    (s) => (window.startDate === null || s.date >= window.startDate) && s.date <= window.endDate,
  );
}
