import { addDaysIso } from '@/core/utils';
import type { BodyField } from '@/domain/body';

import type { BodyAnalytics, DistanceToTarget } from '../bodyAnalytics';
import { insufficient, ok, type MetricResult } from '../metricResult';
import type { NutritionAnalytics, DailyNutrition } from '../nutritionAnalytics';
import { rangeWindow } from '../ranges';
import type { RecompSignal } from '../recomp';
import type { Trend } from '../trend';
import type { WorkoutAnalytics } from '../workoutAnalytics';
import type { InsightContext } from './types';

/** Test builders for the insight engine — a quiet context nothing fires from, plus knobs. */

export const TODAY = '2026-03-15';
const WINDOW = rangeWindow('90d', TODAY);

const none = (): MetricResult<never> => insufficient('no-data', 'x');

export function trend(
  classification: Trend['classification'],
  direction: Trend['direction'],
  slopePerWeek: number,
  spanDays = 90,
): MetricResult<Trend> {
  return ok(
    { classification, direction, slopePerWeek, deltaOverWindow: slopePerWeek * (spanDays / 7) },
    WINDOW,
    { points: 5, spanDays },
  );
}

function quietBody(): BodyAnalytics {
  return {
    weight: { latestKg: 80, trendKg: 80 },
    weightTrend: none(),
    distanceToTarget: none(),
    siteTrends: new Map(),
  };
}

export function distance(over: Partial<DistanceToTarget>): MetricResult<DistanceToTarget> {
  return ok(
    { targetKg: 75, toGoKg: 5, ratePerWeekKg: -0.4, etaWeeks: 12, atGoal: false, ...over },
    WINDOW,
    { points: 1, spanDays: 0 },
  );
}

function quietNutrition(): NutritionAnalytics {
  return {
    daysInRange: 90,
    loggedDays: 0,
    completeness: 0,
    avg: null,
    activeTarget: null,
    calorieAdherence: none(),
    proteinAdherence: none(),
    carbAdherence: none(),
    fatAdherence: none(),
    waterAdherence: none(),
    calorieSkew: null,
    proteinTrend: none(),
  };
}

function quietWorkout(): WorkoutAnalytics {
  return {
    totalWorkouts: 0,
    totalWorkoutsPrev: null,
    frequencyPerWeek: 0,
    consistency: { progress: { completed: 0, planned: 4 }, streak: 0 },
    totalVolumeKg: 0,
    totalVolumePrevKg: null,
    volumeTrend: none(),
    volumeSeries: [],
    volumeByGroup: [],
    pushPull: { ratio: 1, numeratorKg: 100, denominatorKg: 100, flagged: false },
    upperLower: { ratio: 1.5, numeratorKg: 150, denominatorKg: 100, flagged: false },
    mostTrained: null,
    leastTrained: null,
    keyExercises: [],
    avgSessionMinutes: null,
    missedWorkouts: none(),
  };
}

/** Five unremarkable logged days — enough that nothing (good or bad) fires by default. */
function neutralDays(): DailyNutrition[] {
  const protein = [200, 100, 200, 100, 200]; // 3 hits, newest a hit → no miss-streak, not a strong week
  return protein.map((proteinG, i) => ({
    date: addDaysIso(TODAY, -i),
    totals: { kcal: 2500, proteinG, carbG: 250, fatG: 80 }, // kcal/carb/fat on target → no skew
    target: { kcal: 2500, proteinG: 180, carbG: 250, fatG: 80, waterMl: 3000 },
    waterMl: 3000, // water hit → no water-low
    logged: true,
  }));
}

export function baseContext(over: Partial<InsightContext> = {}): InsightContext {
  return {
    today: TODAY,
    window: WINDOW,
    body: quietBody(),
    recomp: none(),
    nutrition: quietNutrition(),
    nutritionDays: neutralDays(),
    workout: quietWorkout(),
    sessions30d: 0,
    completedWeekConsistency: { current: null, previous: null },
    recentPrs: [],
    neglectedGroups: [],
    lastSnapshotDate: TODAY, // recent → housekeeping quiet
    lastPhotoDate: TODAY,
    ...over,
  };
}

export function siteMap(
  entries: readonly [BodyField, MetricResult<Trend>][],
): BodyAnalytics['siteTrends'] {
  return new Map(entries.map(([f, t]) => [f, { latest: 80, trend: t }]));
}

export function recompFired(): MetricResult<RecompSignal> {
  return ok(
    {
      fired: true,
      weightDeltaKg: 0.2,
      waistDeltaCm: -1.5,
      bodyFatDeltaPct: null,
      muscleDeltaKg: null,
      spanDays: 56,
      markers: ['waist -1.5 cm'],
    },
    WINDOW,
    { points: 3, spanDays: 56 },
  );
}

export function loggedDay(over: Partial<DailyNutrition>): DailyNutrition {
  return {
    date: TODAY,
    totals: { kcal: 2000, proteinG: 120, carbG: 200, fatG: 60 },
    target: { kcal: 2500, proteinG: 180, carbG: 250, fatG: 80, waterMl: 3000 },
    waterMl: 0,
    logged: true,
    ...over,
  };
}
