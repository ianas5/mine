import type { EpochMs, IsoDate } from '@/core/utils';
import type { MealSlot, ServingUnit } from '@/domain/nutrition';

/** A reusable food or quick meal (DATABASE §3.5, FITNESS_DOMAIN §4.1). */
export interface Food {
  readonly id: string;
  readonly name: string;
  readonly servingAmount: number;
  readonly servingUnit: ServingUnit;
  readonly kcal: number;
  readonly proteinG: number;
  readonly carbG: number;
  readonly fatG: number;
  readonly isQuickMeal: boolean;
  readonly isCustom: boolean;
  readonly isArchived: boolean;
}

/** A time-versioned set of daily targets (DATABASE §3.5, FITNESS_DOMAIN §4.1). */
export interface NutritionTarget {
  readonly id: string;
  readonly effectiveFrom: IsoDate;
  readonly kcal: number;
  readonly proteinG: number;
  readonly carbG: number;
  readonly fatG: number;
  readonly waterMl: number | null;
}

/** One logged food at a portion, with macros snapshotted at log time (§4.1). */
export interface MealEntry {
  readonly id: string;
  readonly date: IsoDate;
  readonly slot: MealSlot | null;
  /** Provenance to the food (SET NULL on food delete); the snapshot below is authoritative. */
  readonly foodId: string | null;
  readonly foodName: string;
  readonly loggedAmount: number;
  readonly loggedUnit: ServingUnit;
  readonly kcal: number;
  readonly proteinG: number;
  readonly carbG: number;
  readonly fatG: number;
  readonly loggedAt: EpochMs;
}
