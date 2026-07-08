import { sql } from 'drizzle-orm';
import { check, integer, real, sqliteTable } from 'drizzle-orm/sqlite-core';

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
