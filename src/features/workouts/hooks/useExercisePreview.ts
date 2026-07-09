import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import type { ExercisePreview, LoadType } from '@/domain/fitness';

/** Last / Best / Best e1RM for an exercise, refreshed on any workout write. */
export function useExercisePreview(exerciseId: string, loadType: LoadType): ExercisePreview | null {
  const version = useTableVersion('workouts');
  const [preview, setPreview] = useState<ExercisePreview | null>(null);

  useEffect(() => {
    let live = true;
    void workoutRepository.getExercisePreview(exerciseId, loadType).then((p) => {
      if (live) setPreview(p);
    });
    return () => {
      live = false;
    };
  }, [exerciseId, loadType, version]);

  return preview;
}
