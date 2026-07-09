import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import type { Workout } from '@/domain/models';

/** Recent workouts (newest first), refreshed on any workout write. */
export function useRecentWorkouts(): Workout[] | null {
  const version = useTableVersion('workouts');
  const [workouts, setWorkouts] = useState<Workout[] | null>(null);

  useEffect(() => {
    let live = true;
    void workoutRepository.listRecent().then((rows) => {
      if (live) setWorkouts(rows);
    });
    return () => {
      live = false;
    };
  }, [version]);

  return workouts;
}
