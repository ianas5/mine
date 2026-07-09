import { getDb } from '@/core/db';

import { exercises, foods } from '../schema/tables';
import { SEED_EXERCISES } from './exercises';
import { SEED_FOODS } from './foods';

/**
 * Idempotent seeder (DATABASE §5.6): inserts seed exercises and starter foods with
 * stable ids, skipping any that already exist. Re-running never duplicates and never
 * overwrites user edits to seed rows. Runs in the DB-ready gate after migrations.
 */
export async function seedDatabase(): Promise<void> {
  const db = getDb();
  const now = Date.now();

  const exerciseRows = SEED_EXERCISES.map((e) => ({
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
  await db.insert(exercises).values(exerciseRows).onConflictDoNothing({ target: exercises.id });

  const foodRows = SEED_FOODS.map((f) => ({
    id: `food_seed_${f.slug}`,
    name: f.name,
    servingAmount: f.servingAmount,
    servingUnit: f.servingUnit,
    kcal: f.kcal,
    proteinG: f.proteinG,
    carbG: f.carbG,
    fatG: f.fatG,
    isQuickMeal: f.quickMeal === true ? 1 : 0,
    isCustom: 0,
    isArchived: 0,
    createdAt: now,
    updatedAt: now,
  }));
  await db.insert(foods).values(foodRows).onConflictDoNothing({ target: foods.id });
}
