import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import type { NutritionTarget } from '@/domain/models';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';

/** All target eras (newest first) for the targets editor history (UI_UX §4.7). */
export function useTargets(): NutritionTarget[] | undefined {
  const version = useTableVersion('nutrition');
  const [targets, setTargets] = useState<NutritionTarget[] | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void nutritionRepository.listTargets().then((rows) => {
      if (live) setTargets(rows);
    });
    return () => {
      live = false;
    };
  }, [version]);

  return targets;
}
