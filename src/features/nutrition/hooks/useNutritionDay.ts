import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { sumMacros, type MacroSet } from '@/domain/nutrition';
import type { MealEntry } from '@/domain/models';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';

export interface NutritionDay {
  readonly entries: readonly MealEntry[];
  readonly totals: MacroSet;
}

/** A day's meal entries and their totals (FITNESS_DOMAIN §4.2), reactive to writes. */
export function useNutritionDay(date: string): NutritionDay | undefined {
  const version = useTableVersion('nutrition');
  const [day, setDay] = useState<NutritionDay | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void nutritionRepository.listMealEntries(date).then((entries) => {
      if (live) setDay({ entries, totals: sumMacros(entries) });
    });
    return () => {
      live = false;
    };
  }, [date, version]);

  return day;
}
