import { insufficient, ok, type MetricResult, type Range } from './metricResult';
import { linearRegression } from './regression';
import { spanDays, type SeriesPoint } from './timeSeries';

/** §6.4 minimums: a trend needs ≥ 3 points spanning ≥ 14 days. */
export const TREND_MIN_POINTS = 3;
export const TREND_MIN_SPAN_DAYS = 14;

/** Raw sign of movement. */
export type TrendDirection = 'increasing' | 'stable' | 'decreasing';
/**
 * Movement judged through FITNESS_DOMAIN §5.3 directionality. `neutral` = the metric
 * is moving but has no fixed "good" direction (e.g. weight with no goal set) — reported
 * plainly, never coloured good/bad (mirrors the body-comparison `neutral`, Phase 12).
 */
export type TrendClassification = 'improving' | 'stable' | 'declining' | 'neutral';

export interface Trend {
  readonly slopePerWeek: number;
  readonly deltaOverWindow: number;
  readonly classification: TrendClassification;
  readonly direction: TrendDirection;
}

/** Which way is "good" for this metric (§5.3). */
export type GoodDirection = 'lower' | 'higher' | 'neutral';

export interface TrendConfig {
  /** §6.4 stability threshold: |Δ over the window| below this reads as `stable`. */
  readonly stabilityThreshold: number;
  readonly goodDirection: GoodDirection;
  /** Plural noun for the `needed` sentence, e.g. "weigh-ins", "measurements", "sessions". */
  readonly pointNoun: string;
}

function classify(direction: TrendDirection, good: GoodDirection): TrendClassification {
  if (direction === 'stable') return 'stable';
  if (good === 'neutral') return 'neutral';
  const improving = good === 'higher' ? direction === 'increasing' : direction === 'decreasing';
  return improving ? 'improving' : 'declining';
}

/**
 * The canonical trend (FITNESS_DOMAIN §6.4): least-squares regression over the raw
 * in-range points, classified through a stability deadband and §5.3 directionality.
 * Below the minimums it returns `insufficient-data` with a concrete `needed` sentence —
 * never a fabricated line. `points` must already be windowed to `window`.
 */
export function computeTrend(
  points: readonly SeriesPoint[],
  config: TrendConfig,
  window: Range,
): MetricResult<Trend> {
  if (points.length === 0) {
    return insufficient('no-data', `Log some ${config.pointNoun} to see a trend`);
  }
  if (points.length < TREND_MIN_POINTS) {
    const more = TREND_MIN_POINTS - points.length;
    return insufficient(
      'too-few-points',
      `Log ${more} more ${config.pointNoun} (at least ${TREND_MIN_POINTS} across ${TREND_MIN_SPAN_DAYS} days) to see a trend`,
    );
  }
  const span = spanDays(points);
  if (span < TREND_MIN_SPAN_DAYS) {
    return insufficient(
      'span-too-short',
      `Keep logging — a trend needs ${config.pointNoun} spanning at least ${TREND_MIN_SPAN_DAYS} days (currently ${span})`,
    );
  }

  const regression = linearRegression(points);
  if (regression === null) {
    return insufficient(
      'too-few-points',
      `Log ${config.pointNoun} on different days to see a trend`,
    );
  }

  const slopePerDay = regression.slopePerDay;
  const deltaOverWindow = slopePerDay * span;
  const direction: TrendDirection =
    Math.abs(deltaOverWindow) < config.stabilityThreshold
      ? 'stable'
      : slopePerDay > 0
        ? 'increasing'
        : 'decreasing';

  const trend: Trend = {
    slopePerWeek: slopePerDay * 7,
    deltaOverWindow,
    direction,
    classification: classify(direction, config.goodDirection),
  };
  return ok(trend, window, { points: points.length, spanDays: span });
}
