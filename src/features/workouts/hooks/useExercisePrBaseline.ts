import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import {
  computeExerciseBests,
  EMPTY_BESTS,
  type ExerciseBests,
  type LoadType,
} from '@/domain/fitness';

import { useDefaultBodyweight } from './useDefaultBodyweight';

/**
 * The exercise's prior all-time bests (excluding the live session), for the
 * optimistic in-session PR badge (UI_UX §4.1, P15). Recomputed from history — a
 * hint, never an authoritative record. Refreshes on any workout write.
 */
export function useExercisePrBaseline(exerciseId: string, loadType: LoadType): ExerciseBests {
  const version = useTableVersion('workouts');
  const bodyweightKg = useDefaultBodyweight();
  const [bests, setBests] = useState<ExerciseBests>(EMPTY_BESTS);

  useEffect(() => {
    let live = true;
    void workoutRepository.getExerciseSetHistory(exerciseId).then((history) => {
      if (!live) return;
      setBests(
        history === null ? EMPTY_BESTS : computeExerciseBests(history.rows, loadType, bodyweightKg),
      );
    });
    return () => {
      live = false;
    };
  }, [exerciseId, loadType, bodyweightKg, version]);

  return bests;
}
