import { emitTableChanges, getDb, runInTransaction, type TableName } from '@/core/db';
import type { SQLiteTable } from 'drizzle-orm/sqlite-core';

import {
  bodySnapshots,
  exercises,
  foods,
  mealEntries,
  nutritionTargets,
  phases,
  programs,
  progressPhotos,
  sets,
  settings,
  templateExercises,
  templates,
  waterDays,
  workoutExercises,
  workouts,
} from '../schema/tables';
import type { BackupData } from './backupSchema';

/** SQLite caps bound variables per statement; chunk multi-row inserts well under it. */
const INSERT_CHUNK = 100;

const ALL_TABLES: readonly TableName[] = [
  'settings',
  'exercises',
  'workouts',
  'programs',
  'nutrition',
  'body',
  'photos',
  'phases',
];

// Child → parent: safe delete order under foreign_keys = ON (a row is removed only
// after everything referencing it is already gone).
const DELETE_ORDER: readonly SQLiteTable[] = [
  phases,
  progressPhotos,
  bodySnapshots,
  waterDays,
  nutritionTargets,
  mealEntries,
  foods,
  sets,
  workoutExercises,
  workouts,
  templateExercises,
  templates,
  programs,
  exercises,
  settings,
];

async function insertAll(
  table: SQLiteTable,
  rows: readonly Record<string, unknown>[],
): Promise<void> {
  for (let i = 0; i < rows.length; i += INSERT_CHUNK) {
    const chunk = rows.slice(i, i + INSERT_CHUNK);
    await getDb()
      .insert(table)
      .values(chunk as Record<string, unknown>[]);
  }
}

/**
 * Replaces ALL domain data with the imported backup in a single all-or-nothing
 * transaction (DATABASE §6.4). Deletes every table child-first, then inserts every
 * table parent-first (so per-statement FK checks always pass). If any statement
 * throws — a CHECK/FK/UNIQUE violation in a corrupt archive — the transaction rolls
 * back and the prior data is left exactly intact (§6.6). Replace, never merge.
 */
export async function replaceAllData(data: BackupData): Promise<void> {
  await runInTransaction(async () => {
    const db = getDb();
    for (const table of DELETE_ORDER) {
      await db.delete(table);
    }

    // Parent → child (reverse of the delete order), so referenced rows exist first.
    await insertAll(settings, [data.settings]);
    await insertAll(exercises, data.exercises);
    await insertAll(programs, data.programs);
    await insertAll(phases, data.phases);
    await insertAll(templates, data.templates);
    await insertAll(templateExercises, data.templateExercises);
    await insertAll(workouts, data.workouts);
    await insertAll(workoutExercises, data.workoutExercises);
    await insertAll(sets, data.sets);
    await insertAll(foods, data.foods);
    await insertAll(mealEntries, data.mealEntries);
    await insertAll(nutritionTargets, data.nutritionTargets);
    await insertAll(waterDays, data.waterDays);
    await insertAll(bodySnapshots, data.bodySnapshots);
    await insertAll(progressPhotos, data.progressPhotos);
  });

  emitTableChanges(...ALL_TABLES);
}
