export {
  computeExerciseReport,
  EMPTY_REPORT,
  bestE1rmSeries,
  computeExerciseTrend,
  E1RM_STABILITY_KG,
} from './exerciseReport';
export type { ExerciseReport, ExerciseTrend } from './exerciseReport';

export {
  ok,
  insufficient,
  isOk,
  type MetricResult,
  type Range,
  type RangeKey,
  type InsufficientReason,
  type ComputedFrom,
} from './metricResult';
export { RANGE_KEYS, RANGE_LABELS, rangeWindow } from './ranges';
export { sortByDate, pointsInRange, spanDays, type SeriesPoint } from './timeSeries';
export { linearRegression, type Regression } from './regression';
export {
  computeTrend,
  TREND_MIN_POINTS,
  TREND_MIN_SPAN_DAYS,
  type Trend,
  type TrendConfig,
  type TrendDirection,
  type TrendClassification,
  type GoodDirection,
} from './trend';
export { sevenDayMovingAverage, latestMovingAverage } from './movingAverage';
export { bucketSeries, MAX_CHART_POINTS, type Granularity, type Aggregation } from './bucketing';
export {
  computeBodyAnalytics,
  WEIGHT_STABILITY_KG,
  type BodyAnalytics,
  type BodyAnalyticsInput,
  type WeightHeadline,
  type DistanceToTarget,
  type SiteAnalysis,
} from './bodyAnalytics';
