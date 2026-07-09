import { useEffect, useMemo, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { addDaysIso, todayIso } from '@/core/utils';
import { bodyRepository } from '@/data/repositories/bodyRepository';
import { programRepository } from '@/data/repositories/programRepository';
import { settingsRepository } from '@/data/repositories/settingsRepository';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import {
  computeMuscleReports,
  computeWorkoutAnalytics,
  rangeWindow,
  type MuscleGroupReport,
  type RangeKey,
  type TrainingWorkout,
  type WeighIn,
  type WorkoutAnalytics,
} from '@/domain/analytics';

interface Source {
  readonly workouts: readonly TrainingWorkout[];
  readonly weighIns: readonly WeighIn[];
  readonly weeklyWorkoutTarget: number;
  readonly defaultBodyweightKg: number | null;
  readonly scheduledWeekdays: readonly number[];
  readonly firstDate: string | null;
}

export interface TrainingAnalyticsView {
  readonly workout: WorkoutAnalytics;
  readonly muscles: readonly MuscleGroupReport[];
}

/**
 * Training + muscle analytics for a range (ANALYTICS §5.1/§5.6). Repositories fetch the
 * windowed rows; the pure calculators compute; this hook only composes + memoizes by data
 * version and range. `undefined` while first loading.
 */
export function useTrainingAnalytics(range: RangeKey): TrainingAnalyticsView | undefined {
  const version = useTableVersion('workouts', 'programs', 'body', 'settings');
  const [source, setSource] = useState<Source | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void (async () => {
      const today = todayIso();
      const window = rangeWindow(range, today);
      // Fetch enough history for the window + its previous comparison period + the
      // trailing 30-day most/least window; all-time fetches everything.
      const since =
        window.days !== null ? addDaysIso(today, -Math.max(window.days * 2, 30)) : '2000-01-01';

      const [workouts, weighIns, settings, activeProgram] = await Promise.all([
        workoutRepository.getTrainingWorkoutsSince(since),
        bodyRepository.getWeightLog(),
        settingsRepository.get(),
        programRepository.getActiveProgram(),
      ]);

      const scheduledWeekdays = [
        ...new Set(
          (activeProgram?.templates ?? [])
            .map((t) => t.weekday)
            .filter((w): w is number => w !== null),
        ),
      ];

      if (!live) return;
      setSource({
        workouts,
        weighIns,
        weeklyWorkoutTarget: settings.weeklyWorkoutTarget,
        defaultBodyweightKg: settings.defaultBodyweightKg,
        scheduledWeekdays,
        firstDate: workouts[0]?.date ?? null,
      });
    })();
    return () => {
      live = false;
    };
  }, [version, range]);

  return useMemo<TrainingAnalyticsView | undefined>(() => {
    if (source === undefined) return undefined;
    const today = todayIso();
    const window = rangeWindow(range, today, source.firstDate);
    const common = {
      workouts: source.workouts,
      weighIns: source.weighIns,
      defaultBodyweightKg: source.defaultBodyweightKg,
      window,
      today,
    };
    return {
      workout: computeWorkoutAnalytics({
        ...common,
        weeklyWorkoutTarget: source.weeklyWorkoutTarget,
        scheduledWeekdays: source.scheduledWeekdays,
      }),
      muscles: computeMuscleReports(common),
    };
  }, [source, range]);
}
