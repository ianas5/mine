import { sql } from 'drizzle-orm';
import { check, integer, real, sqliteTable, text, uniqueIndex } from 'drizzle-orm/sqlite-core';

import { LOAD_TYPES, MUSCLE_GROUPS } from '@/domain/fitness';

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
