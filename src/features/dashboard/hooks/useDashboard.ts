import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { addDaysIso, daysBetweenIso, isoWeekday, todayIso, type IsoDate } from '@/core/utils';
import { bodyRepository } from '@/data/repositories/bodyRepository';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';
import { programRepository } from '@/data/repositories/programRepository';
import { settingsRepository } from '@/data/repositories/settingsRepository';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import { latestMovingAverage } from '@/domain/analytics';
import {
  currentWeekProgress,
  suggestTemplate,
  weeklyStreak,
  type WeekProgress,
} from '@/domain/fitness';
import { remainingMacros, sumMacros, type MacroSet } from '@/domain/nutrition';
import type { NutritionTarget } from '@/domain/models';

const STREAK_LOOKBACK_DAYS = 84; // 12 weeks — enough context for the streak + recency
const RECENT_TEMPLATE_DAYS = 56; // 8 weeks (UI_UX §5.2)

export type WorkoutCardState = 'planned' | 'done' | 'rest';

export interface DashboardGreeting {
  /** Latest 7-day trend weight, else latest raw weigh-in, else null (greeting line only). */
  readonly trendWeightKg: number | null;
  readonly weighedInToday: boolean;
  /** Days since the last countable workout (null if none in the lookback window). */
  readonly daysSinceLastWorkout: number | null;
}

export interface DashboardWorkout {
  readonly state: WorkoutCardState;
  /** The suggested session's label (template or last-workout name), or null when resting. */
  readonly suggestionLabel: string | null;
}

export interface DashboardNutrition {
  readonly target: NutritionTarget | null;
  readonly totals: MacroSet;
  /** target − consumed per macro; null when no target is active for today. */
  readonly remaining: MacroSet | null;
}

export interface DashboardStreak {
  readonly weeks: number;
  readonly progress: WeekProgress;
}

export interface DashboardData {
  readonly today: IsoDate;
  readonly greeting: DashboardGreeting;
  readonly workout: DashboardWorkout;
  readonly nutrition: DashboardNutrition;
  readonly streak: DashboardStreak;
}

/** Planned sessions for a week: the active program's scheduled templates, else the target. */
function plannedPerWeek(scheduledWeekdayCount: number, weeklyWorkoutTarget: number): number {
  return scheduledWeekdayCount > 0 ? scheduledWeekdayCount : weeklyWorkoutTarget;
}

/**
 * The daily-briefing data (UI_UX §7.2, ANALYTICS §6.5 closed list). Composes nutrition,
 * program suggestion, streak, and trend weight straight from repositories + pure domain
 * — the dashboard is a different feature from workouts/nutrition, so it reads the data
 * layer directly rather than importing their hooks. Reactive to every source table.
 * `undefined` while first loading.
 */
export function useDashboard(): DashboardData | undefined {
  const version = useTableVersion('nutrition', 'workouts', 'programs', 'body', 'settings');
  const [data, setData] = useState<DashboardData | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void (async () => {
      const today = todayIso();
      const weekday = isoWeekday(today);
      const since = addDaysIso(today, -STREAK_LOOKBACK_DAYS);

      const [
        settings,
        entries,
        target,
        activeProgram,
        countableDates,
        recentUses,
        lastList,
        todaySnap,
        snapshots,
      ] = await Promise.all([
        settingsRepository.get(),
        nutritionRepository.listMealEntries(today),
        nutritionRepository.resolveTargetForDate(today),
        programRepository.getActiveProgram(),
        workoutRepository.getCountableWorkoutDatesSince(since),
        programRepository.getRecentTemplateUses(addDaysIso(today, -RECENT_TEMPLATE_DAYS)),
        workoutRepository.listRecent(1),
        bodyRepository.getSnapshot(today),
        bodyRepository.listSnapshots(),
      ]);

      // Nutrition
      const totals = sumMacros(entries);
      const remaining = target ? remainingMacros(target, totals) : null;

      // Today's workout state
      const doneToday = countableDates.includes(today);
      const scheduled = activeProgram?.templates.find((t) => t.weekday === weekday) ?? null;
      const decision = suggestTemplate(
        weekday,
        scheduled ? scheduled.id : null,
        recentUses,
        (lastList[0] ?? null) !== null,
      );
      let suggestionLabel: string | null = null;
      if (decision.kind === 'template') {
        suggestionLabel =
          scheduled?.id === decision.templateId
            ? scheduled.name
            : ((await programRepository.getTemplate(decision.templateId))?.name ?? null);
      } else if (decision.kind === 'repeatLast') {
        suggestionLabel = lastList[0]?.name ?? null;
      }
      const workoutState: WorkoutCardState = doneToday
        ? 'done'
        : suggestionLabel !== null
          ? 'planned'
          : 'rest';

      // Streak (§3.8)
      const scheduledWeekdayCount = (activeProgram?.templates ?? []).filter(
        (t) => t.weekday !== null,
      ).length;
      const planned = plannedPerWeek(scheduledWeekdayCount, settings.weeklyWorkoutTarget);
      const streak = {
        weeks: weeklyStreak(countableDates, today, settings.weeklyWorkoutTarget),
        progress: currentWeekProgress(countableDates, today, planned),
      };

      // Greeting
      const weightPoints = snapshots.flatMap((s) =>
        s.weightKg !== null ? [{ date: s.date, value: s.weightKg }] : [],
      );
      const latestRaw = weightPoints[0]?.value ?? null; // listSnapshots is newest-first
      const lastWorkoutDate = countableDates.length > 0 ? [...countableDates].sort().at(-1)! : null;

      if (!live) return;
      setData({
        today,
        greeting: {
          trendWeightKg: latestMovingAverage(weightPoints) ?? latestRaw,
          weighedInToday: todaySnap?.weightKg != null,
          daysSinceLastWorkout: lastWorkoutDate ? daysBetweenIso(lastWorkoutDate, today) : null,
        },
        workout: { state: workoutState, suggestionLabel },
        nutrition: { target, totals, remaining },
        streak,
      });
    })();
    return () => {
      live = false;
    };
  }, [version]);

  return data;
}
