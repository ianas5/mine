import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { exerciseRepository } from '@/data/repositories/exerciseRepository';
import type { Exercise } from '@/domain/models';

interface UseExercisesResult {
  readonly exercises: Exercise[] | null;
}

interface LoadState {
  readonly key: string;
  readonly exercises: Exercise[];
}

/** Loads active or archived exercises, re-querying on catalog writes (ARCHITECTURE §7). */
export function useExercises(archived: boolean): UseExercisesResult {
  const version = useTableVersion('exercises');
  const key = `${String(archived)}:${version}`;
  const [state, setState] = useState<LoadState | null>(null);

  useEffect(() => {
    let live = true;
    const query = archived ? exerciseRepository.listArchived() : exerciseRepository.listActive();
    void query.then((rows) => {
      if (live) setState({ key, exercises: rows });
    });
    return () => {
      live = false;
    };
  }, [archived, version, key]);

  // Derive the loading state (null) when results for the current key aren't in yet —
  // avoids a synchronous setState in the effect (react-hooks/set-state-in-effect).
  return { exercises: state?.key === key ? state.exercises : null };
}
