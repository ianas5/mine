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

/** `workouts` — completed sessions (DATABASE §3.4). Multiple per date allowed. */
export const workouts = sqliteTable(
  'workouts',
  {
    id: text('id').primaryKey(),
    date: text('date').notNull(),
    name: text('name').notNull(),
    // Provenance only. FK to templates is added in Phase 8 when that table exists (TD-004).
    templateId: text('template_id'),
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
