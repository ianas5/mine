import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import type { Workout } from '@/domain/models';

/** Loads one workout by id, refreshed on workout writes. `undefined` while loading. */
export function useWorkout(id: string): Workout | null | undefined {
  const version = useTableVersion('workouts');
  const [workout, setWorkout] = useState<Workout | null | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void workoutRepository.getById(id).then((w) => {
      if (live) setWorkout(w);
    });
    return () => {
      live = false;
    };
  }, [id, version]);

  return workout;
}
