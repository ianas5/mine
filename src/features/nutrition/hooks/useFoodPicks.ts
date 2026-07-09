import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { nutritionRepository, type FoodPick } from '@/data/repositories/nutritionRepository';

/** Foods for the Log Meal sheet — most-used first, quick meals pinned (UI_UX §4.3). */
export function useFoodPicks(): FoodPick[] | undefined {
  const version = useTableVersion('nutrition');
  const [picks, setPicks] = useState<FoodPick[] | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void nutritionRepository.getFoodPicks().then((rows) => {
      if (live) setPicks(rows);
    });
    return () => {
      live = false;
    };
  }, [version]);

  return picks;
}
