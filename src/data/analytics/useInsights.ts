import { useCallback, useEffect, useMemo, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { insightCooldowns } from '@/core/storage';
import { addDaysIso, todayIso, type IsoDate } from '@/core/utils';
import {
  computeBodyAnalytics,
  computeNutritionAnalytics,
  computeRecompSignal,
  computeWorkoutAnalytics,
  evaluateInsights,
  rangeWindow,
  resolveBodyweightForDate,
  selectDashboardInsights,
  stampCooldowns,
  type Insight,
  type InsightContext,
  type InsightEvidence,
  type TrainingWorkout,
} from '@/domain/analytics';
import type { BodyField } from '@/domain/body';
import {
  MUSCLE_GROUPS,
  countableWorkoutsByWeek,
  detectNewPRs,
  weekConsistencyPercent,
  weekStartIso,
  type ExerciseSetRow,
  type LoadType,
  type MuscleGroup,
} from '@/domain/fitness';

import { bodyRepository } from '../repositories/bodyRepository';
import { nutritionRepository } from '../repositories/nutritionRepository';
import { photoRepository } from '../repositories/photoRepository';
import { programRepository } from '../repositories/programRepository';
import { settingsRepository } from '../repositories/settingsRepository';
import { workoutRepository } from '../repositories/workoutRepository';

const SITE_FIELDS: readonly BodyField[] = [
  'waistCm',
  'chestCm',
  'leftArmCm',
  'rightArmCm',
  'leftThighCm',
  'rightThighCm',
];
const CANONICAL = MUSCLE_GROUPS.filter((g) => g !== 'other');
const PR_WINDOW_DAYS = 7;

export interface InsightsView {
  readonly all: readonly Insight[];
  readonly dashboard: readonly Insight[];
  /** Dismiss an insight = start its cooldown (§6.3); re-evaluates immediately. */
  readonly dismiss: (instanceKey: string, classification: string) => void;
}

interface Assembled {
  readonly context: InsightContext;
}

const isWorkingRow = (loadType: string, warmup: boolean, reps: number, weightKg: number): boolean =>
  !warmup && reps >= 1 && (loadType !== 'external' || weightKg > 0);

/** The route an insight's evidence tap-through opens (UI_UX §8 — one tap from its proof). */
export function insightEvidenceHref(evidence: InsightEvidence): string {
  switch (evidence.kind) {
    case 'muscle-report':
      return '/analytics/muscles';
    case 'measurements':
      return '/measurements';
    case 'photos':
      return '/measurements/photos';
    case 'exercise':
      return `/workouts/exercise/${evidence.exerciseId}`;
    case 'analytics-body':
    case 'analytics-training':
    case 'analytics-nutrition':
      return '/analytics';
  }
}

/** Which primary groups saw ≥ 1 working set since a date (for neglect detection). */
function trainedGroupsSince(
  workouts: readonly TrainingWorkout[],
  since: IsoDate,
): Set<MuscleGroup> {
  const trained = new Set<MuscleGroup>();
  for (const w of workouts) {
    if (w.date < since) continue;
    for (const ex of w.exercises) {
      if (ex.sets.some((s) => isWorkingRow(ex.loadType, s.warmup, s.reps, s.weightKg))) {
        trained.add(ex.primaryMuscleGroup);
      }
    }
  }
  return trained;
}

/** PRs set in the last 7 days: last-week sets vs. all prior history, per exercise (rule 15). */
function recentPrs(
  workouts: readonly TrainingWorkout[],
  today: IsoDate,
  bodyweight: number | null,
): { exerciseId: string; name: string; kinds: string[] }[] {
  const cutoff = addDaysIso(today, -(PR_WINDOW_DAYS - 1));
  const byExercise = new Map<
    string,
    { name: string; loadType: string; prior: ExerciseSetRow[]; candidate: ExerciseSetRow[] }
  >();
  for (const w of workouts) {
    for (const ex of w.exercises) {
      let acc = byExercise.get(ex.exerciseId);
      if (!acc) {
        acc = { name: ex.name, loadType: ex.loadType, prior: [], candidate: [] };
        byExercise.set(ex.exerciseId, acc);
      }
      for (const s of ex.sets) {
        const row: ExerciseSetRow = {
          workoutId: w.id,
          date: w.date,
          workoutOrder: w.startedAt ?? 0,
          weightKg: s.weightKg,
          reps: s.reps,
          warmup: s.warmup,
          counting: ex.counting,
        };
        if (w.date >= cutoff) acc.candidate.push(row);
        else acc.prior.push(row);
      }
    }
  }

  const out: { exerciseId: string; name: string; kinds: string[] }[] = [];
  for (const [exerciseId, acc] of byExercise) {
    if (acc.candidate.length === 0) continue;
    const events = detectNewPRs(acc.prior, acc.candidate, acc.loadType as LoadType, bodyweight);
    if (events.length > 0)
      out.push({ exerciseId, name: acc.name, kinds: events.map((e) => e.kind) });
  }
  return out;
}

/**
 * Assembles the full `InsightContext` from every calculator + raw signals and evaluates the
 * §6.2 rules against the MMKV cooldown map. Shared by the dashboard (top-3) and the Analytics
 * tab (full list); it lives in the data layer because features cannot import one another and
 * both need it (ARCHITECTURE §4). Reactive to every source table; `undefined` while loading.
 */
export function useInsights(): InsightsView | undefined {
  const version = useTableVersion(
    'workouts',
    'programs',
    'body',
    'settings',
    'nutrition',
    'photos',
  );
  const [dismissTick, setDismissTick] = useState(0);
  const [assembled, setAssembled] = useState<Assembled | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void (async () => {
      const today = todayIso();
      const [workouts, weighIns, snapshots, nutritionDays, settings, program, photos] =
        await Promise.all([
          workoutRepository.getTrainingWorkoutsSince('2000-01-01'),
          bodyRepository.getWeightLog(),
          bodyRepository.listSnapshots(),
          nutritionRepository.getDailyNutritionSince(addDaysIso(today, -120)),
          settingsRepository.get(),
          programRepository.getActiveProgram(),
          photoRepository.listPhotos(),
        ]);

      const window90 = rangeWindow('90d', today);
      const scheduledWeekdays = [
        ...new Set(
          (program?.templates ?? []).map((t) => t.weekday).filter((w): w is number => w !== null),
        ),
      ];
      const common = {
        workouts,
        weighIns,
        defaultBodyweightKg: settings.defaultBodyweightKg,
        today,
      };

      const workout90 = computeWorkoutAnalytics({
        ...common,
        window: window90,
        weeklyWorkoutTarget: settings.weeklyWorkoutTarget,
        scheduledWeekdays,
      });
      const workout30 = computeWorkoutAnalytics({
        ...common,
        window: rangeWindow('30d', today),
        weeklyWorkoutTarget: settings.weeklyWorkoutTarget,
        scheduledWeekdays,
      });

      const body = computeBodyAnalytics({
        snapshots,
        window: window90,
        targetWeightKg: settings.targetWeightKg,
        siteFields: SITE_FIELDS,
      });

      // Last two completed weeks' consistency (rules 12/13).
      const countable = [
        ...countableWorkoutsByWeek(
          workouts
            .filter((w) =>
              w.exercises.some((ex) =>
                ex.sets.some((s) => isWorkingRow(ex.loadType, s.warmup, s.reps, s.weightKg)),
              ),
            )
            .map((w) => w.date),
        ).entries(),
      ];
      const weekCount = new Map(countable);
      const planned =
        scheduledWeekdays.length > 0 ? scheduledWeekdays.length : settings.weeklyWorkoutTarget;
      const lastCompleted = addDaysIso(weekStartIso(today), -7);
      const prevCompleted = addDaysIso(weekStartIso(today), -14);
      const completedWeekConsistency = {
        current: weekCount.has(lastCompleted)
          ? weekConsistencyPercent(weekCount.get(lastCompleted)!, planned)
          : null,
        previous: weekCount.has(prevCompleted)
          ? weekConsistencyPercent(weekCount.get(prevCompleted)!, planned)
          : null,
      };

      const bodyweight = resolveBodyweightForDate(weighIns, today, settings.defaultBodyweightKg);
      const trained30 = trainedGroupsSince(workouts, addDaysIso(today, -29));

      const context: InsightContext = {
        today,
        window: window90,
        body,
        recomp: computeRecompSignal(snapshots, today),
        nutrition: computeNutritionAnalytics({ days: nutritionDays, window: window90, today }),
        nutritionDays,
        workout: { ...workout90, pushPull: workout30.pushPull },
        sessions30d: workout30.totalWorkouts,
        completedWeekConsistency,
        recentPrs: recentPrs(workouts, today, bodyweight),
        neglectedGroups: CANONICAL.filter((g) => !trained30.has(g)),
        lastSnapshotDate: snapshots[0]?.date ?? null,
        lastPhotoDate: photos[0]?.date ?? null,
      };

      if (live) setAssembled({ context });
    })();
    return () => {
      live = false;
    };
  }, [version]);

  const dismiss = useCallback((instanceKey: string, classification: string) => {
    const next = stampCooldowns(
      [{ instanceKey, classification } as Insight],
      insightCooldowns.get(),
      todayIso(),
    );
    insightCooldowns.set(next);
    setDismissTick((t) => t + 1); // trigger a re-evaluation that now suppresses it
  }, []);

  return useMemo<InsightsView | undefined>(() => {
    if (assembled === undefined) return undefined;
    void dismissTick; // re-evaluate after a dismiss stamps a cooldown
    const all = evaluateInsights(assembled.context, insightCooldowns.get());
    return { all, dashboard: selectDashboardInsights(all), dismiss };
  }, [assembled, dismissTick, dismiss]);
}
