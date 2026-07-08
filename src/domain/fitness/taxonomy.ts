/** Muscle taxonomy & load types — canonical vocabulary from FITNESS_DOMAIN §3.3 / §3.4. */

/** The 11 canonical muscle groups plus the `other` fallback, in display order. */
export const MUSCLE_GROUPS = [
  'chest',
  'shoulders',
  'back',
  'biceps',
  'triceps',
  'forearms',
  'core',
  'glutes',
  'quads',
  'hamstrings',
  'calves',
  'other',
] as const;
export type MuscleGroup = (typeof MUSCLE_GROUPS)[number];

export const MUSCLE_GROUP_LABELS: Record<MuscleGroup, string> = {
  chest: 'Chest',
  shoulders: 'Shoulders',
  back: 'Back',
  biceps: 'Biceps',
  triceps: 'Triceps',
  forearms: 'Forearms',
  core: 'Core',
  glutes: 'Glutes',
  quads: 'Quads',
  hamstrings: 'Hamstrings',
  calves: 'Calves',
  other: 'Other',
};

/** Load types define how a set's weight is interpreted (FITNESS_DOMAIN §3.4). */
export const LOAD_TYPES = [
  'external',
  'bodyweight',
  'bodyweight_plus',
  'assisted',
  'timed',
] as const;
export type LoadType = (typeof LOAD_TYPES)[number];

export const LOAD_TYPE_LABELS: Record<LoadType, string> = {
  external: 'External load',
  bodyweight: 'Bodyweight',
  bodyweight_plus: 'Bodyweight + added',
  assisted: 'Assisted',
  timed: 'Timed hold',
};

export type MovementPattern = 'push' | 'pull' | 'legs' | 'core';

/** Push/Pull/Legs classification derived from primary muscle group (FITNESS_DOMAIN §3.3). */
export const MOVEMENT_PATTERN: Record<MuscleGroup, MovementPattern | null> = {
  chest: 'push',
  shoulders: 'push',
  triceps: 'push',
  back: 'pull',
  biceps: 'pull',
  forearms: 'pull',
  glutes: 'legs',
  quads: 'legs',
  hamstrings: 'legs',
  calves: 'legs',
  core: 'core',
  other: null,
};

export type BodySplit = 'upper' | 'lower' | 'core';

/** Upper/Lower split (FITNESS_DOMAIN §3.3); core counted separately. */
export const BODY_SPLIT: Record<MuscleGroup, BodySplit | null> = {
  chest: 'upper',
  shoulders: 'upper',
  back: 'upper',
  biceps: 'upper',
  triceps: 'upper',
  forearms: 'upper',
  glutes: 'lower',
  quads: 'lower',
  hamstrings: 'lower',
  calves: 'lower',
  core: 'core',
  other: null,
};
