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
  resolveBodyweightForDate,
  type TrainingSet,
  type TrainingExercise,
  type TrainingWorkout,
  type WeighIn,
} from './trainingData';
export {
  computeWorkoutAnalytics,
  PUSH_PULL_BAND,
  UPPER_LOWER_BAND,
  type WorkoutAnalytics,
  type WorkoutAnalyticsInput,
  type Balance,
  type GroupVolume,
  type KeyExerciseStrength,
} from './workoutAnalytics';
export {
  computeMuscleReports,
  reportForGroup,
  type MuscleGroupReport,
  type MuscleAnalyticsInput,
  type ExerciseRef,
} from './muscleAnalytics';
export { computeRecompSignal, type RecompSignal } from './recomp';
export {
  computePhaseReport,
  PHASE_MIN_DAYS,
  PHASE_MIN_SNAPSHOTS,
  type PhaseReport,
  type PhaseReportInput,
  type PhaseBodyDeltas,
  type PhaseTrainingSummary,
  type PhaseNutritionSummary,
  type PhaseIntentVerdict,
  type PhaseWeeklyRates,
  type PhasePr,
  type IntentAlignment,
} from './phaseAnalytics';
export {
  evaluateInsights,
  selectDashboardInsights,
  stampCooldowns,
  RULES,
  type Insight,
  type InsightContext,
  type InsightCategory,
  type InsightTone,
  type InsightEvidence,
  type CooldownMap,
} from './insights';
export {
  computeNutritionAnalytics,
  proteinMissStreak,
  trailingSignals,
  calorieSkew,
  type NutritionAnalytics,
  type NutritionAnalyticsInput,
  type DailyNutrition,
  type AdherenceStat,
} from './nutritionAnalytics';
export {
  computeBodyAnalytics,
  WEIGHT_STABILITY_KG,
  type BodyAnalytics,
  type BodyAnalyticsInput,
  type WeightHeadline,
  type DistanceToTarget,
  type SiteAnalysis,
} from './bodyAnalytics';
