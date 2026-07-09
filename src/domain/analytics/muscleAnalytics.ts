import { addDaysIso, daysBetweenIso, type IsoDate } from '@/core/utils';
import {
  MUSCLE_GROUPS,
  computeExerciseBests,
  isWorkingSet,
  setVolumeKg,
  weekStartIso,
  type ExerciseSetRow,
  type LoadType,
  type MuscleGroup,
} from '@/domain/fitness';

import { computeExerciseTrend } from './exerciseReport';
import { type MetricResult, type Range } from './metricResult';
import { computeTrend, type Trend } from './trend';
import type { SeriesPoint } from './timeSeries';
import { resolveBodyweightForDate, type TrainingWorkout, type WeighIn } from './trainingData';

const REPORT_WINDOW_DAYS = 30;

export interface ExerciseRef {
  readonly exerciseId: string;
  readonly name: string;
  /** e1RM (kg) for `strongest`; e1RM slope (kg/week) for the improving refs. */
  readonly value: number;
}

export interface MuscleGroupReport {
  readonly group: MuscleGroup;
  readonly volume30dKg: number;
  readonly workingSets30d: number;
  readonly currentWeekVolumeKg: number;
  /** Sessions touching the group per week over the window. */
  readonly frequencyPerWeek: number;
  readonly strongest: ExerciseRef | null;
  readonly fastestImproving: ExerciseRef | null;
  readonly weakestImproving: ExerciseRef | null;
  readonly volumeTrend: MetricResult<Trend>;
  readonly lastTrained: { readonly date: IsoDate; readonly daysSince: number } | null;
  /** True when the group has no working sets in the window — the honest zero state. */
  readonly untrained: boolean;
}

export interface MuscleAnalyticsInput {
  readonly workouts: readonly TrainingWorkout[];
  readonly weighIns: readonly WeighIn[];
  readonly defaultBodyweightKg: number | null;
  readonly window: Range;
  readonly today: IsoDate;
}

interface ExerciseAcc {
  name: string;
  loadType: LoadType;
  rows: ExerciseSetRow[];
}

/** Reports for every canonical muscle group (never-trained groups return the zero state). */
export function computeMuscleReports(input: MuscleAnalyticsInput): MuscleGroupReport[] {
  return MUSCLE_GROUPS.filter((g) => g !== 'other').map((group) => reportForGroup(input, group));
}

/**
 * One muscle group's report (ANALYTICS §5.6), pure — a coach's read of the last months:
 * volume/sets/frequency, the **strongest** lift, the **fastest-** and **weakest-improving**
 * lifts (by e1RM slope), the volume trend, and when it was last trained. A group with no
 * working sets in the window returns an honest zero state, never fabricated numbers.
 */
export function reportForGroup(input: MuscleAnalyticsInput, group: MuscleGroup): MuscleGroupReport {
  const { workouts, weighIns, defaultBodyweightKg: fallback, window, today } = input;
  const windowed = workouts.filter(
    (w) => (window.startDate === null || w.date >= window.startDate) && w.date <= window.endDate,
  );

  const weekStartOfToday = weekStartIso(today);
  const since30 = addDaysIso(today, -(REPORT_WINDOW_DAYS - 1));

  const byExercise = new Map<string, ExerciseAcc>();
  const sessions = new Set<string>();
  const weeklyVolume = new Map<IsoDate, number>();
  let volume30dKg = 0;
  let workingSets30d = 0;
  let currentWeekVolumeKg = 0;
  let lastDate: IsoDate | null = null;

  for (const workout of windowed) {
    const bw = resolveBodyweightForDate(weighIns, workout.date, fallback);
    let touched = false;

    for (const exercise of workout.exercises) {
      if (exercise.primaryMuscleGroup !== group) continue;

      let acc = byExercise.get(exercise.exerciseId);
      if (!acc) {
        acc = { name: exercise.name, loadType: exercise.loadType, rows: [] };
        byExercise.set(exercise.exerciseId, acc);
      }

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
        if (!isWorkingSet(set, exercise.loadType)) continue;
        touched = true;
        const vol = setVolumeKg(set, exercise.loadType, bw, exercise.counting);
        weeklyVolume.set(
          weekStartIso(workout.date),
          (weeklyVolume.get(weekStartIso(workout.date)) ?? 0) + vol,
        );
        if (workout.date >= since30) {
          volume30dKg += vol;
          workingSets30d += 1;
        }
        if (weekStartIso(workout.date) === weekStartOfToday) currentWeekVolumeKg += vol;
      }
    }

    if (touched) {
      sessions.add(workout.id);
      if (lastDate === null || workout.date > lastDate) lastDate = workout.date;
    }
  }

  const untrained = sessions.size === 0;
  const bodyweight = resolveBodyweightForDate(weighIns, today, fallback);

  // Strongest + improving refs across the group's exercises.
  let strongest: ExerciseRef | null = null;
  let fastest: ExerciseRef | null = null;
  let weakest: ExerciseRef | null = null;
  for (const [exerciseId, acc] of byExercise) {
    const bestE1rm = computeExerciseBests(acc.rows, acc.loadType, bodyweight).bestE1rmKg;
    if (bestE1rm !== null && (strongest === null || bestE1rm > strongest.value)) {
      strongest = { exerciseId, name: acc.name, value: bestE1rm };
    }
    const trend = computeExerciseTrend(acc.rows, acc.loadType, bodyweight, today).trend;
    if (trend.status === 'ok') {
      const slope = trend.value.slopePerWeek;
      if (fastest === null || slope > fastest.value) {
        fastest = { exerciseId, name: acc.name, value: slope };
      }
      if (weakest === null || slope < weakest.value) {
        weakest = { exerciseId, name: acc.name, value: slope };
      }
    }
  }

  const rangeDays =
    window.days ?? Math.max(1, daysBetweenIso(windowed[0]?.date ?? today, today) + 1);
  const volumeSeries: SeriesPoint[] = [...weeklyVolume.entries()]
    .map(([date, value]) => ({ date, value }))
    .sort((a, b) => (a.date < b.date ? -1 : 1));
  const totalGroupVolume = volumeSeries.reduce((s, p) => s + p.value, 0);

  return {
    group,
    volume30dKg,
    workingSets30d,
    currentWeekVolumeKg,
    frequencyPerWeek: sessions.size / (rangeDays / 7),
    strongest,
    fastestImproving: fastest,
    weakestImproving: weakest,
    volumeTrend: computeTrend(
      volumeSeries,
      {
        stabilityThreshold: Math.max(1, totalGroupVolume * 0.05),
        goodDirection: 'neutral',
        pointNoun: 'weeks',
      },
      window,
    ),
    lastTrained: lastDate ? { date: lastDate, daysSince: daysBetweenIso(lastDate, today) } : null,
    untrained,
  };
}
