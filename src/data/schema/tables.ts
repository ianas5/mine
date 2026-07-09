import { sql } from 'drizzle-orm';
import {
  check,
  index,
  integer,
  real,
  sqliteTable,
  text,
  uniqueIndex,
} from 'drizzle-orm/sqlite-core';

import { LOAD_TYPES, MUSCLE_GROUPS } from '@/domain/fitness';
import { MEAL_SLOTS, SERVING_UNITS } from '@/domain/nutrition';

const UNILATERAL_COUNTING = ['none', 'single_doubled', 'per_side'] as const;

// Inline the enum literals into an `IN`-style CHECK. Values are trusted domain
// constants (not user input); sql.raw is required because string params would
// otherwise serialize as `?` bind placeholders, which CHECK cannot use.
const oneOf = (columnName: string, values: readonly string[]) =>
  sql.raw(values.map((v) => `${columnName} = '${v}'`).join(' OR '));

/**
 * `settings` — single-row domain configuration (DATABASE §3.1).
 * One row (id = 1), created by the seeder; upsert-only; included in backups.
 */
export const settings = sqliteTable(
  'settings',
  {
    id: integer('id').primaryKey(),
    weeklyWorkoutTarget: integer('weekly_workout_target').notNull().default(4),
    defaultBodyweightKg: real('default_bodyweight_kg'),
    heightCm: real('height_cm'),
    waterCupMl: integer('water_cup_ml').notNull().default(250),
    createdAt: integer('created_at').notNull(),
    updatedAt: integer('updated_at').notNull(),
  },
  (table) => [check('settings_single_row', sql`${table.id} = 1`)],
);

export type SettingsRow = typeof settings.$inferSelect;

/**
 * `exercises` — the catalog (DATABASE §3.2). Seed rows use stable ids
 * `ex_seed_<slug>`; case-insensitive unique name; referenced rows are archived,
 * not deleted (enforced by the repository + FK RESTRICT once history exists).
 */
export const exercises = sqliteTable(
  'exercises',
  {
    id: text('id').primaryKey(),
    name: text('name').notNull(),
    primaryMuscleGroup: text('primary_muscle_group').notNull(),
    secondaryMuscleGroups: text('secondary_muscle_groups').notNull().default('[]'),
    loadType: text('load_type').notNull().default('external'),
    defaultUnilateral: integer('default_unilateral').notNull().default(0),
    isCustom: integer('is_custom').notNull().default(0),
    isArchived: integer('is_archived').notNull().default(0),
    notes: text('notes'),
    createdAt: integer('created_at').notNull(),
    updatedAt: integer('updated_at').notNull(),
  },
  (table) => [
    uniqueIndex('exercises_name_nocase').on(sql`${table.name} COLLATE NOCASE`),
    check('exercises_primary_group', oneOf('primary_muscle_group', MUSCLE_GROUPS)),
    check('exercises_load_type', oneOf('load_type', LOAD_TYPES)),
  ],
);

export type ExerciseRow = typeof exercises.$inferSelect;

/**
 * `programs` — a named collection of session templates (DATABASE §3.3). At most one
 * is active; that invariant is enforced in the repository transaction, not the schema.
 */
export const programs = sqliteTable('programs', {
  id: text('id').primaryKey(),
  name: text('name').notNull(),
  notes: text('notes'),
  isActive: integer('is_active').notNull().default(0),
  isArchived: integer('is_archived').notNull().default(0),
  createdAt: integer('created_at').notNull(),
  updatedAt: integer('updated_at').notNull(),
});
export type ProgramRow = typeof programs.$inferSelect;

/** `templates` — a single-session blueprint (DATABASE §3.3). NULL program = standalone. */
export const templates = sqliteTable(
  'templates',
  {
    id: text('id').primaryKey(),
    programId: text('program_id').references(() => programs.id, { onDelete: 'cascade' }),
    name: text('name').notNull(),
    position: integer('position').notNull().default(0),
    // 0 = Monday … 6 = Sunday; enables the weekday-scheduled suggestion (§3.8/§5.2).
    weekday: integer('weekday'),
    notes: text('notes'),
    isArchived: integer('is_archived').notNull().default(0),
    createdAt: integer('created_at').notNull(),
    updatedAt: integer('updated_at').notNull(),
  },
  (table) => [
    index('templates_program').on(table.programId),
    check(
      'templates_weekday',
      sql`${table.weekday} IS NULL OR (${table.weekday} >= 0 AND ${table.weekday} <= 6)`,
    ),
  ],
);
export type TemplateRow = typeof templates.$inferSelect;

/** `template_exercises` — planned targets, never performance (DATABASE §3.3). */
export const templateExercises = sqliteTable(
  'template_exercises',
  {
    id: text('id').primaryKey(),
    templateId: text('template_id')
      .notNull()
      .references(() => templates.id, { onDelete: 'cascade' }),
    exerciseId: text('exercise_id')
      .notNull()
      .references(() => exercises.id, { onDelete: 'restrict' }),
    position: integer('position').notNull(),
    targetSets: integer('target_sets'),
    targetRepMin: integer('target_rep_min'),
    targetRepMax: integer('target_rep_max'),
    targetRpe: real('target_rpe'),
    restSeconds: integer('rest_seconds'),
    notes: text('notes'),
  },
  (table) => [index('template_exercises_template').on(table.templateId, table.position)],
);
export type TemplateExerciseRow = typeof templateExercises.$inferSelect;

/** `workouts` — completed sessions (DATABASE §3.4). Multiple per date allowed. */
export const workouts = sqliteTable(
  'workouts',
  {
    id: text('id').primaryKey(),
    date: text('date').notNull(),
    name: text('name').notNull(),
    // Provenance only (which template started this session). SET NULL on template
    // delete — a workout is history and is never rewritten by template changes.
    templateId: text('template_id').references(() => templates.id, { onDelete: 'set null' }),
    startedAt: integer('started_at'),
    endedAt: integer('ended_at'),
    notes: text('notes'),
    createdAt: integer('created_at').notNull(),
    updatedAt: integer('updated_at').notNull(),
  },
  (table) => [index('workouts_date').on(table.date)],
);
export type WorkoutRow = typeof workouts.$inferSelect;

/** `workout_exercises` — one exercise as performed within a workout (DATABASE §3.4). */
export const workoutExercises = sqliteTable(
  'workout_exercises',
  {
    id: text('id').primaryKey(),
    workoutId: text('workout_id')
      .notNull()
      .references(() => workouts.id, { onDelete: 'cascade' }),
    exerciseId: text('exercise_id')
      .notNull()
      .references(() => exercises.id, { onDelete: 'restrict' }),
    position: integer('position').notNull(),
    unilateralCounting: text('unilateral_counting').notNull().default('none'),
    notes: text('notes'),
  },
  (table) => [
    index('workout_exercises_workout').on(table.workoutId),
    index('workout_exercises_exercise').on(table.exerciseId),
    check('workout_exercises_counting', oneOf('unilateral_counting', UNILATERAL_COUNTING)),
  ],
);
export type WorkoutExerciseRow = typeof workoutExercises.$inferSelect;

/**
 * `workout_drafts` — crash-safe active-session checkpoint (DATABASE §3.4,
 * ARCHITECTURE §7.1). Deliberately a single-row JSON blob (`id = 1`): a recovery
 * snapshot of the in-memory session, rewritten on debounce/background and deleted
 * on finish/discard. Zod-validated before resume; excluded from every backup path.
 */
export const workoutDrafts = sqliteTable(
  'workout_drafts',
  {
    id: integer('id').primaryKey(),
    payload: text('payload').notNull(),
    updatedAt: integer('updated_at').notNull(),
  },
  (table) => [check('workout_drafts_single_row', sql`${table.id} = 1`)],
);
export type WorkoutDraftRow = typeof workoutDrafts.$inferSelect;

/** `sets` — one performed set (DATABASE §3.4). `reps` holds seconds for timed exercises. */
export const sets = sqliteTable(
  'sets',
  {
    id: text('id').primaryKey(),
    workoutExerciseId: text('workout_exercise_id')
      .notNull()
      .references(() => workoutExercises.id, { onDelete: 'cascade' }),
    position: integer('position').notNull(),
    weightKg: real('weight_kg').notNull().default(0),
    reps: integer('reps').notNull().default(0),
    rpe: real('rpe'),
    rir: integer('rir'),
    isWarmup: integer('is_warmup').notNull().default(0),
    notes: text('notes'),
  },
  (table) => [
    index('sets_workout_exercise').on(table.workoutExerciseId),
    check(
      'sets_rpe_range',
      sql`${table.rpe} IS NULL OR (${table.rpe} >= 0 AND ${table.rpe} <= 10)`,
    ),
    check(
      'sets_rir_range',
      sql`${table.rir} IS NULL OR (${table.rir} >= 0 AND ${table.rir} <= 10)`,
    ),
  ],
);
export type SetRow = typeof sets.$inferSelect;

/**
 * `foods` — reusable foods and quick meals (DATABASE §3.5, FITNESS_DOMAIN §4.1).
 * Seed rows use stable ids `food_seed_<slug>`. No fiber column (FITNESS_DOMAIN §4).
 */
export const foods = sqliteTable(
  'foods',
  {
    id: text('id').primaryKey(),
    name: text('name').notNull(),
    servingAmount: real('serving_amount').notNull(),
    servingUnit: text('serving_unit').notNull(),
    kcal: integer('kcal').notNull(),
    proteinG: real('protein_g').notNull(),
    carbG: real('carb_g').notNull(),
    fatG: real('fat_g').notNull(),
    isQuickMeal: integer('is_quick_meal').notNull().default(0),
    isCustom: integer('is_custom').notNull().default(1),
    isArchived: integer('is_archived').notNull().default(0),
    createdAt: integer('created_at').notNull(),
    updatedAt: integer('updated_at').notNull(),
  },
  (table) => [check('foods_serving_unit', oneOf('serving_unit', SERVING_UNITS))],
);
export type FoodRow = typeof foods.$inferSelect;

/**
 * `meal_entries` — the log; **macros snapshotted at log time** (DATABASE §3.5,
 * FITNESS_DOMAIN §4.1/§4.2). Editing a food never rewrites past entries; day totals
 * are `SUM … GROUP BY date`, never a stored aggregate. `food_id` is SET NULL on
 * delete (provenance) so history survives a food deletion.
 */
export const mealEntries = sqliteTable(
  'meal_entries',
  {
    id: text('id').primaryKey(),
    date: text('date').notNull(),
    slot: text('slot'),
    foodId: text('food_id').references(() => foods.id, { onDelete: 'set null' }),
    foodName: text('food_name').notNull(),
    loggedAmount: real('logged_amount').notNull(),
    loggedUnit: text('logged_unit').notNull(),
    kcal: integer('kcal').notNull(),
    proteinG: real('protein_g').notNull(),
    carbG: real('carb_g').notNull(),
    fatG: real('fat_g').notNull(),
    loggedAt: integer('logged_at').notNull(),
  },
  (table) => [
    index('meal_entries_date').on(table.date),
    index('meal_entries_food').on(table.foodId),
    check('meal_entries_slot', sql`${table.slot} IS NULL OR (${oneOf('slot', MEAL_SLOTS)})`),
  ],
);
export type MealEntryRow = typeof mealEntries.$inferSelect;

/**
 * `nutrition_targets` — time-versioned daily goals (DATABASE §3.5, FITNESS_DOMAIN
 * §4.1). The active target for a date is the row with the greatest
 * `effective_from ≤ date`; resolution is implemented **once** in
 * `nutritionRepository`. The UNIQUE index doubles as the lookup index.
 */
export const nutritionTargets = sqliteTable('nutrition_targets', {
  id: text('id').primaryKey(),
  effectiveFrom: text('effective_from').notNull().unique(),
  kcal: integer('kcal').notNull(),
  proteinG: real('protein_g').notNull(),
  carbG: real('carb_g').notNull(),
  fatG: real('fat_g').notNull(),
  waterMl: integer('water_ml'),
  createdAt: integer('created_at').notNull(),
  updatedAt: integer('updated_at').notNull(),
});
export type NutritionTargetRow = typeof nutritionTargets.$inferSelect;

/**
 * `water_days` — cumulative daily water (DATABASE §3.5, §2.5 daily-keyed). An
 * absent row is *unlogged*; a row with `ml = 0` is a real logged zero.
 */
export const waterDays = sqliteTable('water_days', {
  date: text('date').primaryKey(),
  ml: integer('ml').notNull().default(0),
  updatedAt: integer('updated_at').notNull(),
});
export type WaterDayRow = typeof waterDays.$inferSelect;

/**
 * `body_snapshots` — one per date, field-merged (DATABASE §3.6, FITNESS_DOMAIN
 * §5.1). Every field is optional; bilateral sites are stored per side, never
 * collapsed. Saving updates only the fields present in the input; clearing a field
 * is a separate, deliberate action (the repository enforces the merge contract).
 */
export const bodySnapshots = sqliteTable('body_snapshots', {
  date: text('date').primaryKey(),
  weightKg: real('weight_kg'),
  bodyFatPct: real('body_fat_pct'),
  muscleMassKg: real('muscle_mass_kg'),
  visceralFat: real('visceral_fat'),
  bmi: real('bmi'),
  neckCm: real('neck_cm'),
  chestCm: real('chest_cm'),
  waistCm: real('waist_cm'),
  hipsCm: real('hips_cm'),
  leftArmCm: real('left_arm_cm'),
  rightArmCm: real('right_arm_cm'),
  leftForearmCm: real('left_forearm_cm'),
  rightForearmCm: real('right_forearm_cm'),
  leftThighCm: real('left_thigh_cm'),
  rightThighCm: real('right_thigh_cm'),
  leftCalfCm: real('left_calf_cm'),
  rightCalfCm: real('right_calf_cm'),
  createdAt: integer('created_at').notNull(),
  updatedAt: integer('updated_at').notNull(),
});
export type BodySnapshotRow = typeof bodySnapshots.$inferSelect;
