import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import type { Food } from '@/domain/models';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';

/** One food by id, reactive. `undefined` while loading, `null` when missing. */
export function useFood(id: string): Food | null | undefined {
  const version = useTableVersion('nutrition');
  const [food, setFood] = useState<Food | null | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void nutritionRepository.getFood(id).then((row) => {
      if (live) setFood(row);
    });
    return () => {
      live = false;
    };
  }, [id, version]);

  return food;
}
