import { and, asc, eq } from 'drizzle-orm';

import { emitTableChanges, getDb } from '@/core/db';
import type { LoadType, MuscleGroup } from '@/domain/fitness';
import type { Exercise } from '@/domain/models';

import { newId } from '../id';
import { rowToExercise } from '../mappers/exerciseMapper';
import { exercises } from '../schema/tables';

export interface CreateExerciseInput {
  readonly name: string;
  readonly primaryMuscleGroup: MuscleGroup;
  readonly loadType: LoadType;
  readonly defaultUnilateral: boolean;
  readonly notes?: string | null;
}

/** Thrown when a custom exercise name collides case-insensitively (DATABASE §3.2). */
export class DuplicateExerciseNameError extends Error {
  constructor(name: string) {
    super(`An exercise named "${name}" already exists.`);
    this.name = 'DuplicateExerciseNameError';
  }
}

/**
 * Exercise catalog (DATABASE §3.2/§7). Referenced rows are archived, not deleted;
 * hard delete is confined to custom, unreferenced rows (no referencing tables
 * exist until Phase 4, at which point FK RESTRICT enforces it at the DB level).
 */
export const exerciseRepository = {
  async listActive(): Promise<Exercise[]> {
    const rows = await getDb()
      .select()
      .from(exercises)
      .where(eq(exercises.isArchived, 0))
      .orderBy(asc(exercises.name));
    return rows.map(rowToExercise);
  },

  async listArchived(): Promise<Exercise[]> {
    const rows = await getDb()
      .select()
      .from(exercises)
      .where(eq(exercises.isArchived, 1))
      .orderBy(asc(exercises.name));
    return rows.map(rowToExercise);
  },

  async createCustom(input: CreateExerciseInput): Promise<Exercise> {
    const db = getDb();
    const id = newId('ex_custom');
    const now = Date.now();
    try {
      await db.insert(exercises).values({
        id,
        name: input.name.trim(),
        primaryMuscleGroup: input.primaryMuscleGroup,
        secondaryMuscleGroups: '[]',
        loadType: input.loadType,
        defaultUnilateral: input.defaultUnilateral ? 1 : 0,
        isCustom: 1,
        isArchived: 0,
        notes: input.notes ?? null,
        createdAt: now,
        updatedAt: now,
      });
    } catch (cause) {
      // Read the driver message/code WITHOUT gating on `instanceof Error`: the
      // SQLite driver's error type binds `Error` to whichever module realm loaded
      // the native addon first, so `instanceof` is unreliable across Jest's
      // per-file realms and would let a genuine UNIQUE violation escape unmapped.
      const detail =
        typeof cause === 'object' && cause !== null && 'message' in cause
          ? String((cause as { message: unknown }).message)
          : String(cause);
      const code =
        typeof cause === 'object' && cause !== null && 'code' in cause
          ? String((cause as { code: unknown }).code)
          : '';
      if (/UNIQUE|constraint/i.test(detail) || code.startsWith('SQLITE_CONSTRAINT')) {
        throw new DuplicateExerciseNameError(input.name.trim());
      }
      throw cause;
    }
    emitTableChanges('exercises');
    const rows = await db.select().from(exercises).where(eq(exercises.id, id));
    const created = rows[0];
    if (!created) {
      throw new Error('custom exercise could not be created');
    }
    return rowToExercise(created);
  },

  async archive(id: string): Promise<void> {
    await getDb()
      .update(exercises)
      .set({ isArchived: 1, updatedAt: Date.now() })
      .where(eq(exercises.id, id));
    emitTableChanges('exercises');
  },

  async unarchive(id: string): Promise<void> {
    await getDb()
      .update(exercises)
      .set({ isArchived: 0, updatedAt: Date.now() })
      .where(eq(exercises.id, id));
    emitTableChanges('exercises');
  },

  /** Hard-deletes a CUSTOM exercise only; seed rows can only be archived (§3.2). */
  async remove(id: string): Promise<void> {
    await getDb()
      .delete(exercises)
      .where(and(eq(exercises.id, id), eq(exercises.isCustom, 1)));
    emitTableChanges('exercises');
  },
} as const;
