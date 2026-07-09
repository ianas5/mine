import { getDb } from '@/core/db';

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

/**
 * Reads every backed-up table into a `BackupData` block (DATABASE §6 export path).
 * `data/backup` is inside the persistence boundary and may use Drizzle directly for
 * the full dump — the repositories-only rule binds *features*, not this service.
 * `workout_drafts` is deliberately excluded (ephemeral recovery state, §3.4).
 */
export async function collectBackupData(): Promise<BackupData> {
  const db = getDb();
  const settingsRows = await db.select().from(settings);
  const settingsRow = settingsRows[0];
  if (!settingsRow) {
    // The seeder guarantees the single settings row; its absence means a broken DB.
    throw new Error('Cannot export: the settings row is missing.');
  }

  return {
    settings: settingsRow,
    exercises: await db.select().from(exercises),
    programs: await db.select().from(programs),
    phases: await db.select().from(phases),
    templates: await db.select().from(templates),
    templateExercises: await db.select().from(templateExercises),
    workouts: await db.select().from(workouts),
    workoutExercises: await db.select().from(workoutExercises),
    sets: await db.select().from(sets),
    foods: await db.select().from(foods),
    mealEntries: await db.select().from(mealEntries),
    nutritionTargets: await db.select().from(nutritionTargets),
    waterDays: await db.select().from(waterDays),
    bodySnapshots: await db.select().from(bodySnapshots),
    progressPhotos: await db.select().from(progressPhotos),
  };
}
