import { getDb } from '@/core/db';

import { exercises } from '../schema/tables';
import { SEED_EXERCISES } from './exercises';

/**
 * Idempotent seeder (DATABASE §5.6): inserts seed exercises with stable ids,
 * skipping any that already exist. Re-running never duplicates and never
 * overwrites user edits to seed rows. Runs in the DB-ready gate after migrations.
 */
export async function seedDatabase(): Promise<void> {
  const db = getDb();
  const now = Date.now();
  const rows = SEED_EXERCISES.map((e) => ({
    id: `ex_seed_${e.slug}`,
    name: e.name,
    primaryMuscleGroup: e.group,
    secondaryMuscleGroups: '[]',
    loadType: e.loadType ?? 'external',
    defaultUnilateral: e.unilateral === true ? 1 : 0,
    isCustom: 0,
    isArchived: 0,
    notes: null,
    createdAt: now,
    updatedAt: now,
  }));
  // onConflictDoNothing on the primary key → insert-if-absent semantics.
  await db.insert(exercises).values(rows).onConflictDoNothing({ target: exercises.id });
}
