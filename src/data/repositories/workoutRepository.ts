import { and, asc, eq } from 'drizzle-orm';

import { emitTableChanges, getDb, runInTransaction } from '@/core/db';
import { todayIso } from '@/core/utils';
import type { LoadType, UnilateralCounting } from '@/domain/fitness';
import type { Workout, WorkoutExercise } from '@/domain/models';

import { newId } from '../id';
import { exercises, sets, workoutExercises, workouts } from '../schema/tables';

export interface NewSetInput {
  readonly weightKg: number;
  readonly reps: number;
  readonly rpe: number | null;
  readonly warmup: boolean;
}

export interface NewWorkoutExerciseInput {
  readonly exerciseId: string;
  readonly unilateralCounting: UnilateralCounting;
  readonly notes: string | null;
  readonly sets: readonly NewSetInput[];
}

export interface NewWorkoutInput {
  readonly name: string;
  readonly startedAt: number | null;
  readonly endedAt: number | null;
  readonly notes: string | null;
  readonly exercises: readonly NewWorkoutExerciseInput[];
}

/** A set is worth persisting if it was performed (reps) or is a warm-up marker. */
function isPersistableSet(set: NewSetInput): boolean {
  return set.reps > 0 || set.warmup;
}

export const workoutRepository = {
  /**
   * Persists a completed session in one all-or-nothing transaction (DATABASE §7).
   * Empty sets and exercises with no persistable sets are dropped. Returns the id.
   */
  async saveCompletedWorkout(input: NewWorkoutInput): Promise<string> {
    const db = getDb();
    const now = Date.now();
    const workoutId = newId('wk');

    const entries = input.exercises
      .map((exercise) => ({ exercise, keptSets: exercise.sets.filter(isPersistableSet) }))
      .filter((entry) => entry.keptSets.length > 0);

    await runInTransaction(async () => {
      await db.insert(workouts).values({
        id: workoutId,
        date: todayIso(),
        name: input.name.trim() || 'Workout',
        templateId: null,
        startedAt: input.startedAt,
        endedAt: input.endedAt,
        notes: input.notes,
        createdAt: now,
        updatedAt: now,
      });

      for (const [exIndex, { exercise, keptSets }] of entries.entries()) {
        const workoutExerciseId = newId('we');
        await db.insert(workoutExercises).values({
          id: workoutExerciseId,
          workoutId,
          exerciseId: exercise.exerciseId,
          position: exIndex,
          unilateralCounting: exercise.unilateralCounting,
          notes: exercise.notes,
        });
        for (const [setIndex, set] of keptSets.entries()) {
          await db.insert(sets).values({
            id: newId('set'),
            workoutExerciseId,
            position: setIndex,
            weightKg: set.weightKg,
            reps: set.reps,
            rpe: set.rpe,
            rir: null,
            isWarmup: set.warmup ? 1 : 0,
            notes: null,
          });
        }
      }
    });

    emitTableChanges('workouts');
    return workoutId;
  },

  /** Loads a full workout tree (joins exercise name/load type) or null. */
  async getById(id: string): Promise<Workout | null> {
    const db = getDb();
    const workoutRows = await db.select().from(workouts).where(eq(workouts.id, id));
    const workout = workoutRows[0];
    if (!workout) return null;

    const exerciseRows = await db
      .select({
        id: workoutExercises.id,
        exerciseId: workoutExercises.exerciseId,
        name: exercises.name,
        loadType: exercises.loadType,
        unilateralCounting: workoutExercises.unilateralCounting,
        position: workoutExercises.position,
        notes: workoutExercises.notes,
      })
      .from(workoutExercises)
      .innerJoin(exercises, eq(workoutExercises.exerciseId, exercises.id))
      .where(eq(workoutExercises.workoutId, id))
      .orderBy(asc(workoutExercises.position));

    const built: WorkoutExercise[] = [];
    for (const ex of exerciseRows) {
      const setRows = await db
        .select()
        .from(sets)
        .where(eq(sets.workoutExerciseId, ex.id))
        .orderBy(asc(sets.position));
      built.push({
        id: ex.id,
        exerciseId: ex.exerciseId,
        name: ex.name,
        loadType: ex.loadType as LoadType,
        unilateralCounting: ex.unilateralCounting as UnilateralCounting,
        position: ex.position,
        notes: ex.notes,
        sets: setRows.map((s) => ({
          id: s.id,
          position: s.position,
          weightKg: s.weightKg,
          reps: s.reps,
          rpe: s.rpe,
          rir: s.rir,
          warmup: s.isWarmup === 1,
          notes: s.notes,
        })),
      });
    }

    return {
      id: workout.id,
      date: workout.date,
      name: workout.name,
      startedAt: workout.startedAt,
      endedAt: workout.endedAt,
      notes: workout.notes,
      exercises: built,
    };
  },

  /** Count of workouts on a given date (supports the two-per-day rule, edge 3). */
  async countOnDate(date: string): Promise<number> {
    const rows = await getDb()
      .select({ id: workouts.id })
      .from(workouts)
      .where(eq(workouts.date, date));
    return rows.length;
  },

  /** Deletes a workout and its cascade tree. */
  async remove(id: string): Promise<void> {
    await getDb()
      .delete(workouts)
      .where(and(eq(workouts.id, id)));
    emitTableChanges('workouts');
  },
} as const;
