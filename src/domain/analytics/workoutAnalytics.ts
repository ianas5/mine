import { addDaysIso, daysBetweenIso, isoWeekday, type IsoDate } from '@/core/utils';
import {
  BODY_SPLIT,
  MOVEMENT_PATTERN,
  MUSCLE_GROUPS,
  currentWeekProgress,
  isWorkingSet,
  setVolumeKg,
  weekStartIso,
  weeklyStreak,
  type ExerciseSetRow,
  type MuscleGroup,
  type WeekProgress,
} from '@/domain/fitness';

import { computeExerciseTrend } from './exerciseReport';
import { insufficient, ok, type MetricResult, type Range } from './metricResult';
import { computeTrend, type Trend } from './trend';
import type { SeriesPoint } from './timeSeries';
import { resolveBodyweightForDate, type TrainingWorkout, type WeighIn } from './trainingData';

/** Balance bands (ANALYTICS §5.1): push:pull healthy 0.8–1.25; upper:lower 1.0–2.0. */
export const PUSH_PULL_BAND = { low: 0.8, high: 1.25 } as const;
export const UPPER_LOWER_BAND = { low: 1.0, high: 2.0 } as const;
const MOST_LEAST_WINDOW_DAYS = 30;
const KEY_EXERCISE_COUNT = 5;

export interface Balance {
  /** numerator ÷ denominator, or null when the denominator is 0 (can't compare). */
  readonly ratio: number | null;
  readonly numeratorKg: number;
  readonly denominatorKg: number;
  /** Outside the healthy band → surfaced softly as attention. */
  readonly flagged: boolean;
}

export interface GroupVolume {
  readonly group: MuscleGroup;
  readonly volumeKg: number;
  readonly workingSets: number;
}

export interface KeyExerciseStrength {
  readonly exerciseId: string;
  readonly name: string;
  readonly sessions: number;
  readonly trend: MetricResult<Trend>;
}

export interface WorkoutAnalytics {
  readonly totalWorkouts: number;
  readonly totalWorkoutsPrev: number | null;
  readonly frequencyPerWeek: number;
  readonly consistency: { readonly progress: WeekProgress; readonly streak: number };
  readonly totalVolumeKg: number;
  readonly totalVolumePrevKg: number | null;
  readonly volumeTrend: MetricResult<Trend>;
  /** Weekly volume buckets for the chart (already downsampled by week). */
  readonly volumeSeries: readonly SeriesPoint[];
  /** Non-zero groups, ranked by volume desc. */
  readonly volumeByGroup: readonly GroupVolume[];
  readonly pushPull: Balance;
  readonly upperLower: Balance;
  readonly mostTrained: GroupVolume | null;
  readonly leastTrained: GroupVolume | null;
  readonly keyExercises: readonly KeyExerciseStrength[];
  readonly avgSessionMinutes: number | null;
  /** Count of missed scheduled sessions; `no-target-set` without a weekday schedule. */
  readonly missedWorkouts: MetricResult<number>;
}

export interface WorkoutAnalyticsInput {
  /** Workouts over a broad window (≥ current range + the preceding comparison period). */
  readonly workouts: readonly TrainingWorkout[];
  readonly weighIns: readonly WeighIn[];
  readonly weeklyWorkoutTarget: number;
  readonly defaultBodyweightKg: number | null;
  /** The active program's scheduled weekdays (0 = Mon); empty ⇒ no schedule. */
  readonly scheduledWeekdays: readonly number[];
  readonly window: Range;
  readonly today: IsoDate;
}

const inWindow = (
  workouts: readonly TrainingWorkout[],
  start: IsoDate | null,
  end: IsoDate,
): TrainingWorkout[] =>
  workouts.filter((w) => (start === null || w.date >= start) && w.date <= end);

/** A working set's volume load with the workout-date bodyweight resolved (§3.4/§3.5). */
function workingSetVolume(
  workout: TrainingWorkout,
  weighIns: readonly WeighIn[],
  fallback: number | null,
): { volumeKg: number; byGroup: Map<MuscleGroup, GroupVolume>; workingSets: number } {
  const bw = resolveBodyweightForDate(weighIns, workout.date, fallback);
  const byGroup = new Map<MuscleGroup, GroupVolume>();
  let volumeKg = 0;
  let workingSets = 0;

  for (const exercise of workout.exercises) {
    let exVolume = 0;
    let exSets = 0;
    for (const set of exercise.sets) {
      if (!isWorkingSet(set, exercise.loadType)) continue;
      exSets += 1;
      exVolume += setVolumeKg(set, exercise.loadType, bw, exercise.counting);
    }
    if (exSets === 0) continue;
    volumeKg += exVolume;
    workingSets += exSets;
    const group = exercise.primaryMuscleGroup;
    const prior = byGroup.get(group);
    byGroup.set(group, {
      group,
      volumeKg: (prior?.volumeKg ?? 0) + exVolume,
      workingSets: (prior?.workingSets ?? 0) + exSets,
    });
  }
  return { volumeKg, byGroup, workingSets };
}

function aggregate(
  workouts: readonly TrainingWorkout[],
  weighIns: readonly WeighIn[],
  fallback: number | null,
): { totalVolumeKg: number; byGroup: Map<MuscleGroup, GroupVolume>; countable: IsoDate[] } {
  const byGroup = new Map<MuscleGroup, GroupVolume>();
  const countable: IsoDate[] = [];
  let totalVolumeKg = 0;

  for (const workout of workouts) {
    const w = workingSetVolume(workout, weighIns, fallback);
    if (w.workingSets === 0) continue; // not countable
    countable.push(workout.date);
    totalVolumeKg += w.volumeKg;
    for (const [group, gv] of w.byGroup) {
      const prior = byGroup.get(group);
      byGroup.set(group, {
        group,
        volumeKg: (prior?.volumeKg ?? 0) + gv.volumeKg,
        workingSets: (prior?.workingSets ?? 0) + gv.workingSets,
      });
    }
  }
  return { totalVolumeKg, byGroup, countable };
}

function balance(
  byGroup: Map<MuscleGroup, GroupVolume>,
  pick: (group: MuscleGroup) => 'num' | 'den' | null,
  band: { low: number; high: number },
): Balance {
  let numeratorKg = 0;
  let denominatorKg = 0;
  for (const gv of byGroup.values()) {
    const side = pick(gv.group);
    if (side === 'num') numeratorKg += gv.volumeKg;
    else if (side === 'den') denominatorKg += gv.volumeKg;
  }
  const ratio = denominatorKg > 0 ? numeratorKg / denominatorKg : null;
  const flagged = ratio !== null && (ratio < band.low || ratio > band.high);
  return { ratio, numeratorKg, denominatorKg, flagged };
}

function weeklyVolumeSeries(
  workouts: readonly TrainingWorkout[],
  weighIns: readonly WeighIn[],
  fallback: number | null,
): SeriesPoint[] {
  const byWeek = new Map<IsoDate, number>();
  for (const workout of workouts) {
    const { volumeKg } = workingSetVolume(workout, weighIns, fallback);
    if (volumeKg <= 0) continue;
    const week = weekStartIso(workout.date);
    byWeek.set(week, (byWeek.get(week) ?? 0) + volumeKg);
  }
  return [...byWeek.entries()]
    .map(([date, value]) => ({ date, value }))
    .sort((a, b) => (a.date < b.date ? -1 : 1));
}

function keyExerciseStrength(
  windowed: readonly TrainingWorkout[],
  weighIns: readonly WeighIn[],
  fallback: number | null,
  today: IsoDate,
): KeyExerciseStrength[] {
  interface Acc {
    name: string;
    loadType: TrainingWorkout['exercises'][number]['loadType'];
    sessions: Set<string>;
    rows: ExerciseSetRow[];
  }
  const byExercise = new Map<string, Acc>();
  for (const workout of windowed) {
    for (const exercise of workout.exercises) {
      let acc = byExercise.get(exercise.exerciseId);
      if (!acc) {
        acc = { name: exercise.name, loadType: exercise.loadType, sessions: new Set(), rows: [] };
        byExercise.set(exercise.exerciseId, acc);
      }
      acc.sessions.add(workout.id);
      for (const set of exercise.sets) {
        acc.rows.push({
          workoutId: workout.id,
          date: workout.date,
          workoutOrder: 0,
          weightKg: set.weightKg,
          reps: set.reps,
          warmup: set.warmup,
          counting: exercise.counting,
        });
      }
    }
  }

  const bodyweight = resolveBodyweightForDate(weighIns, today, fallback);
  return [...byExercise.entries()]
    .map(([exerciseId, acc]) => ({ exerciseId, acc }))
    .sort((a, b) => b.acc.sessions.size - a.acc.sessions.size)
    .slice(0, KEY_EXERCISE_COUNT)
    .map(({ exerciseId, acc }) => ({
      exerciseId,
      name: acc.name,
      sessions: acc.sessions.size,
      trend: computeExerciseTrend(acc.rows, acc.loadType, bodyweight, today).trend,
    }));
}

function missedWorkouts(
  input: WorkoutAnalyticsInput,
  countableInWindow: readonly IsoDate[],
): MetricResult<number> {
  if (input.scheduledWeekdays.length === 0) {
    return insufficient(
      'no-target-set',
      'Add a weekday schedule to a program to track missed sessions',
    );
  }
  const countableDays = new Set(countableInWindow);
  const start = input.window.startDate ?? input.workouts[0]?.date ?? input.today;
  let missed = 0;
  for (let date = start; date < input.today; date = addDaysIso(date, 1)) {
    if (input.scheduledWeekdays.includes(isoWeekday(date)) && !countableDays.has(date)) missed += 1;
  }
  return ok(missed, input.window, { points: missed, spanDays: daysBetweenIso(start, input.today) });
}

/**
 * The Workout analytics calculator (ANALYTICS §5.1), pure. Answers the training-quality
 * questions — consistency, strength progression (key exercises), balance (push:pull,
 * upper:lower), neglect (most/least trained) — alongside volume as context. Every
 * data-gated metric is a `MetricResult`; `other`-group volume is excluded from balance
 * (§3.3); unilateral doubling comes from the stored marker only (§3.5).
 */
export function computeWorkoutAnalytics(input: WorkoutAnalyticsInput): WorkoutAnalytics {
  const { workouts, weighIns, defaultBodyweightKg: fallback, window, today } = input;

  const windowed = inWindow(workouts, window.startDate, window.endDate);
  const cur = aggregate(windowed, weighIns, fallback);

  // Previous equal-length period (no comparison for all-time).
  let prev: ReturnType<typeof aggregate> | null = null;
  if (window.days !== null && window.startDate !== null) {
    const prevEnd = addDaysIso(window.startDate, -1);
    const prevStart = addDaysIso(window.startDate, -window.days);
    prev = aggregate(inWindow(workouts, prevStart, prevEnd), weighIns, fallback);
  }

  const rangeDays =
    window.days ?? Math.max(1, daysBetweenIso(cur.countable[0] ?? today, today) + 1);
  const frequencyPerWeek = cur.countable.length / (rangeDays / 7);

  const volumeSeries = weeklyVolumeSeries(windowed, weighIns, fallback);
  const volumeTrend = computeTrend(
    volumeSeries,
    {
      stabilityThreshold: Math.max(1, cur.totalVolumeKg * 0.05),
      goodDirection: 'neutral', // more volume isn't universally "better" (P: quality > quantity)
      pointNoun: 'weeks of training',
    },
    window,
  );

  // Most/least trained over the trailing 30 days, every canonical group, zeros included.
  const last30 = aggregate(
    inWindow(workouts, addDaysIso(today, -(MOST_LEAST_WINDOW_DAYS - 1)), today),
    weighIns,
    fallback,
  );
  const canonical = MUSCLE_GROUPS.filter((g) => g !== 'other');
  const trained30: GroupVolume[] = canonical.map(
    (group) => last30.byGroup.get(group) ?? { group, volumeKg: 0, workingSets: 0 },
  );
  const mostTrained = trained30.reduce<GroupVolume | null>(
    (best, gv) => (best === null || gv.workingSets > best.workingSets ? gv : best),
    null,
  );
  const leastTrained = trained30.reduce<GroupVolume | null>(
    (worst, gv) => (worst === null || gv.workingSets < worst.workingSets ? gv : worst),
    null,
  );

  const durations = windowed
    .filter((w) => w.startedAt !== null && w.endedAt !== null)
    .map((w) => (w.endedAt! - w.startedAt!) / 60_000);
  const avgSessionMinutes =
    durations.length > 0 ? durations.reduce((s, m) => s + m, 0) / durations.length : null;

  return {
    totalWorkouts: cur.countable.length,
    totalWorkoutsPrev: prev ? prev.countable.length : null,
    frequencyPerWeek,
    consistency: {
      progress: currentWeekProgress(cur.countable, today, plannedPerWeek(input)),
      streak: weeklyStreak(cur.countable, today, input.weeklyWorkoutTarget),
    },
    totalVolumeKg: cur.totalVolumeKg,
    totalVolumePrevKg: prev ? prev.totalVolumeKg : null,
    volumeTrend,
    volumeSeries,
    volumeByGroup: [...cur.byGroup.values()].sort((a, b) => b.volumeKg - a.volumeKg),
    pushPull: balance(cur.byGroup, (g) => patternSide(g, 'push', 'pull'), PUSH_PULL_BAND),
    upperLower: balance(cur.byGroup, (g) => splitSide(g), UPPER_LOWER_BAND),
    mostTrained,
    leastTrained,
    keyExercises: keyExerciseStrength(windowed, weighIns, fallback, today),
    avgSessionMinutes,
    missedWorkouts: missedWorkouts(input, cur.countable),
  };
}

function plannedPerWeek(input: WorkoutAnalyticsInput): number {
  return input.scheduledWeekdays.length > 0
    ? input.scheduledWeekdays.length
    : input.weeklyWorkoutTarget;
}

function patternSide(
  group: MuscleGroup,
  num: 'push' | 'pull',
  den: 'push' | 'pull',
): 'num' | 'den' | null {
  const pattern = MOVEMENT_PATTERN[group];
  if (pattern === num) return 'num';
  if (pattern === den) return 'den';
  return null; // legs, core, other excluded from push:pull
}

function splitSide(group: MuscleGroup): 'num' | 'den' | null {
  const split = BODY_SPLIT[group];
  if (split === 'upper') return 'num';
  if (split === 'lower') return 'den';
  return null; // core, other excluded from upper:lower
}
