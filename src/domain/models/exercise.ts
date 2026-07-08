import type { LoadType, MuscleGroup } from '@/domain/fitness';

/** A catalog exercise definition (FITNESS_DOMAIN §3.1, DATABASE §3.2). */
export interface Exercise {
  readonly id: string;
  readonly name: string;
  readonly primaryMuscleGroup: MuscleGroup;
  /** Stored for v2; NOT credited in v1 volume math (FITNESS_DOMAIN §3.3). */
  readonly secondaryMuscleGroups: readonly MuscleGroup[];
  readonly loadType: LoadType;
  /** Prefills the per-entry unilateral marker when logging (FITNESS_DOMAIN §3.4). */
  readonly defaultUnilateral: boolean;
  readonly isCustom: boolean;
  readonly isArchived: boolean;
  readonly notes: string | null;
}
