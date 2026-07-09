import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import {
  dayAdherence,
  remainingMacros,
  sumMacros,
  type DayAdherence,
  type MacroSet,
} from '@/domain/nutrition';
import type { MealEntry, NutritionTarget } from '@/domain/models';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';

export interface NutritionDay {
  readonly entries: readonly MealEntry[];
  readonly totals: MacroSet;
  readonly target: NutritionTarget | null;
  /** Logged water in ml, or null when the day is unlogged (0 ≠ absent). */
  readonly waterMl: number | null;
  /** target − consumed per macro; null when no target is active for the date. */
  readonly remaining: MacroSet | null;
  /** Per-macro adherence; null when no target is active (insufficient-data). */
  readonly adherence: DayAdherence | null;
}

/**
 * A day's entries, totals, resolved target, water, remaining, and adherence
 * (FITNESS_DOMAIN §4.2/§4.3). The target is resolved through the single canonical
 * path (`resolveTargetForDate`); this hook never re-derives it from dates.
 */
export function useNutritionDay(date: string): NutritionDay | undefined {
  const version = useTableVersion('nutrition');
  const [day, setDay] = useState<NutritionDay | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void Promise.all([
      nutritionRepository.listMealEntries(date),
      nutritionRepository.resolveTargetForDate(date),
      nutritionRepository.getWater(date),
    ]).then(([entries, target, waterMl]) => {
      if (!live) return;
      const totals = sumMacros(entries);
      setDay({
        entries,
        totals,
        target,
        waterMl,
        remaining: target ? remainingMacros(target, totals) : null,
        adherence: target ? dayAdherence(target, totals, waterMl ?? 0) : null,
      });
    });
    return () => {
      live = false;
    };
  }, [date, version]);

  return day;
}
