import type { LoadType, MuscleGroup } from '@/domain/fitness';

/**
 * Seed exercise library (DATABASE §5.6) — machine/cable-rich hypertrophy catalog
 * matching the user profile. Stable ids `ex_seed_<slug>`; insert-if-absent, so
 * user edits are never overwritten and new items ship as new slugs.
 */
export interface SeedExercise {
  readonly slug: string;
  readonly name: string;
  readonly group: MuscleGroup;
  readonly loadType?: LoadType;
  readonly unilateral?: boolean;
}

const E = (
  slug: string,
  name: string,
  group: MuscleGroup,
  loadType?: LoadType,
  unilateral?: boolean,
): SeedExercise => ({
  slug,
  name,
  group,
  ...(loadType ? { loadType } : {}),
  ...(unilateral ? { unilateral } : {}),
});

export const SEED_EXERCISES: readonly SeedExercise[] = [
  // Chest
  E('barbell-bench-press', 'Barbell Bench Press', 'chest'),
  E('incline-barbell-bench-press', 'Incline Barbell Bench Press', 'chest'),
  E('decline-barbell-bench-press', 'Decline Barbell Bench Press', 'chest'),
  E('dumbbell-bench-press', 'Dumbbell Bench Press', 'chest'),
  E('incline-dumbbell-press', 'Incline Dumbbell Press', 'chest'),
  E('machine-chest-press', 'Machine Chest Press', 'chest'),
  E('incline-machine-press', 'Incline Machine Press', 'chest'),
  E('pec-deck', 'Pec Deck', 'chest'),
  E('cable-fly', 'Cable Fly', 'chest'),
  E('low-to-high-cable-fly', 'Low-to-High Cable Fly', 'chest'),
  E('high-to-low-cable-fly', 'High-to-Low Cable Fly', 'chest'),
  E('dumbbell-fly', 'Dumbbell Fly', 'chest'),
  E('push-up', 'Push-Up', 'chest', 'bodyweight'),
  E('chest-dip', 'Chest Dip', 'chest', 'bodyweight'),

  // Shoulders
  E('overhead-barbell-press', 'Overhead Barbell Press', 'shoulders'),
  E('seated-dumbbell-shoulder-press', 'Seated Dumbbell Shoulder Press', 'shoulders'),
  E('machine-shoulder-press', 'Machine Shoulder Press', 'shoulders'),
  E('arnold-press', 'Arnold Press', 'shoulders'),
  E('dumbbell-lateral-raise', 'Dumbbell Lateral Raise', 'shoulders'),
  E('cable-lateral-raise', 'Cable Lateral Raise', 'shoulders', 'external', true),
  E('machine-lateral-raise', 'Machine Lateral Raise', 'shoulders'),
  E('dumbbell-front-raise', 'Dumbbell Front Raise', 'shoulders'),
  E('cable-front-raise', 'Cable Front Raise', 'shoulders'),
  E('reverse-pec-deck', 'Reverse Pec Deck', 'shoulders'),
  E('cable-rear-delt-fly', 'Cable Rear Delt Fly', 'shoulders'),
  E('face-pull', 'Face Pull', 'shoulders'),
  E('barbell-upright-row', 'Barbell Upright Row', 'shoulders'),
  E('dumbbell-shrug', 'Dumbbell Shrug', 'shoulders'),
  E('cable-shrug', 'Cable Shrug', 'shoulders'),

  // Back
  E('deadlift', 'Deadlift', 'back'),
  E('lat-pulldown', 'Lat Pulldown', 'back'),
  E('wide-grip-lat-pulldown', 'Wide-Grip Lat Pulldown', 'back'),
  E('close-grip-lat-pulldown', 'Close-Grip Lat Pulldown', 'back'),
  E('pull-up', 'Pull-Up', 'back', 'bodyweight'),
  E('chin-up', 'Chin-Up', 'back', 'bodyweight'),
  E('assisted-pull-up', 'Assisted Pull-Up', 'back', 'assisted'),
  E('seated-cable-row', 'Seated Cable Row', 'back'),
  E('barbell-row', 'Barbell Row', 'back'),
  E('pendlay-row', 'Pendlay Row', 'back'),
  E('dumbbell-row', 'Dumbbell Row', 'back', 'external', true),
  E('chest-supported-row', 'Chest-Supported Row', 'back'),
  E('t-bar-row', 'T-Bar Row', 'back'),
  E('machine-row', 'Machine Row', 'back'),
  E('straight-arm-pulldown', 'Straight-Arm Pulldown', 'back'),
  E('machine-pullover', 'Machine Pullover', 'back'),

  // Biceps
  E('barbell-curl', 'Barbell Curl', 'biceps'),
  E('ez-bar-curl', 'EZ-Bar Curl', 'biceps'),
  E('dumbbell-curl', 'Dumbbell Curl', 'biceps'),
  E('alternating-dumbbell-curl', 'Alternating Dumbbell Curl', 'biceps', 'external', true),
  E('hammer-curl', 'Hammer Curl', 'biceps'),
  E('incline-dumbbell-curl', 'Incline Dumbbell Curl', 'biceps'),
  E('preacher-curl', 'Preacher Curl', 'biceps'),
  E('cable-curl', 'Cable Curl', 'biceps'),
  E('cable-rope-hammer-curl', 'Cable Rope Hammer Curl', 'biceps'),
  E('concentration-curl', 'Concentration Curl', 'biceps', 'external', true),
  E('spider-curl', 'Spider Curl', 'biceps'),

  // Triceps
  E('cable-tricep-pushdown', 'Cable Tricep Pushdown', 'triceps'),
  E('rope-pushdown', 'Rope Pushdown', 'triceps'),
  E('overhead-cable-extension', 'Overhead Cable Extension', 'triceps'),
  E('overhead-dumbbell-extension', 'Overhead Dumbbell Extension', 'triceps'),
  E('skull-crusher', 'Skull Crusher', 'triceps'),
  E('close-grip-bench-press', 'Close-Grip Bench Press', 'triceps'),
  E('tricep-dip', 'Tricep Dip', 'triceps', 'bodyweight'),
  E('bench-dip', 'Bench Dip', 'triceps', 'bodyweight'),
  E('machine-tricep-extension', 'Machine Tricep Extension', 'triceps'),
  E('cable-tricep-kickback', 'Cable Tricep Kickback', 'triceps', 'external', true),

  // Forearms
  E('barbell-wrist-curl', 'Barbell Wrist Curl', 'forearms'),
  E('reverse-wrist-curl', 'Reverse Wrist Curl', 'forearms'),
  E('cable-wrist-curl', 'Cable Wrist Curl', 'forearms'),
  E('reverse-barbell-curl', 'Reverse Barbell Curl', 'forearms'),
  E('farmers-carry', "Farmer's Carry", 'forearms'),

  // Core
  E('plank', 'Plank', 'core', 'timed'),
  E('side-plank', 'Side Plank', 'core', 'timed', true),
  E('hanging-leg-raise', 'Hanging Leg Raise', 'core', 'bodyweight'),
  E('hanging-knee-raise', 'Hanging Knee Raise', 'core', 'bodyweight'),
  E('cable-crunch', 'Cable Crunch', 'core'),
  E('machine-crunch', 'Machine Crunch', 'core'),
  E('crunch', 'Crunch', 'core', 'bodyweight'),
  E('russian-twist', 'Russian Twist', 'core', 'bodyweight'),
  E('ab-wheel-rollout', 'Ab Wheel Rollout', 'core', 'bodyweight'),
  E('cable-woodchop', 'Cable Woodchop', 'core', 'external', true),
  E('lying-leg-raise', 'Lying Leg Raise', 'core', 'bodyweight'),
  E('mountain-climbers', 'Mountain Climbers', 'core', 'bodyweight'),

  // Glutes
  E('barbell-hip-thrust', 'Barbell Hip Thrust', 'glutes'),
  E('machine-hip-thrust', 'Machine Hip Thrust', 'glutes'),
  E('glute-bridge', 'Glute Bridge', 'glutes'),
  E('bulgarian-split-squat', 'Bulgarian Split Squat', 'glutes', 'external', true),
  E('cable-glute-kickback', 'Cable Glute Kickback', 'glutes', 'external', true),
  E('machine-glute-kickback', 'Machine Glute Kickback', 'glutes', 'external', true),
  E('step-up', 'Step-Up', 'glutes', 'external', true),
  E('sumo-deadlift', 'Sumo Deadlift', 'glutes'),

  // Quads
  E('back-squat', 'Back Squat', 'quads'),
  E('front-squat', 'Front Squat', 'quads'),
  E('hack-squat', 'Hack Squat', 'quads'),
  E('leg-press', 'Leg Press', 'quads'),
  E('leg-extension', 'Leg Extension', 'quads'),
  E('goblet-squat', 'Goblet Squat', 'quads'),
  E('walking-lunge', 'Walking Lunge', 'quads', 'external', true),
  E('smith-machine-squat', 'Smith Machine Squat', 'quads'),

  // Hamstrings
  E('romanian-deadlift', 'Romanian Deadlift', 'hamstrings'),
  E('lying-leg-curl', 'Lying Leg Curl', 'hamstrings'),
  E('seated-leg-curl', 'Seated Leg Curl', 'hamstrings'),
  E('stiff-leg-deadlift', 'Stiff-Leg Deadlift', 'hamstrings'),
  E('good-morning', 'Good Morning', 'hamstrings'),
  E('nordic-curl', 'Nordic Curl', 'hamstrings', 'bodyweight'),
  E('cable-pull-through', 'Cable Pull-Through', 'hamstrings'),

  // Calves
  E('standing-calf-raise', 'Standing Calf Raise', 'calves'),
  E('seated-calf-raise', 'Seated Calf Raise', 'calves'),
  E('leg-press-calf-raise', 'Leg Press Calf Raise', 'calves'),
  E('smith-machine-calf-raise', 'Smith Machine Calf Raise', 'calves'),
];
