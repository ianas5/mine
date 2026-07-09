import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { todayIso } from '@/core/utils';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import {
  computeExerciseReport,
  computeExerciseTrend,
  type ExerciseReport,
  type ExerciseTrend,
} from '@/domain/analytics';

import { useDefaultBodyweight } from './useDefaultBodyweight';

export interface ExerciseReportView {
  readonly name: string;
  readonly report: ExerciseReport;
  /** e1RM strength trend + sparkline series (ANALYTICS §5.1/§5.5). */
  readonly trend: ExerciseTrend;
}

/**
 * The all-time Exercise Report (ANALYTICS §5.5), recomputed from history on every
 * workout write so edits and deletes make records recede (never cached). Returns
 * `undefined` while loading and `null` when the exercise does not exist.
 */
export function useExerciseReport(exerciseId: string): ExerciseReportView | null | undefined {
  const version = useTableVersion('workouts');
  const bodyweightKg = useDefaultBodyweight();
  const [view, setView] = useState<ExerciseReportView | null | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void workoutRepository.getExerciseSetHistory(exerciseId).then((history) => {
      if (!live) return;
      setView(
        history === null
          ? null
          : {
              name: history.name,
              report: computeExerciseReport(history.rows, history.loadType, bodyweightKg),
              trend: computeExerciseTrend(history.rows, history.loadType, bodyweightKg, todayIso()),
            },
      );
    });
    return () => {
      live = false;
    };
  }, [exerciseId, bodyweightKg, version]);

  return view;
}
