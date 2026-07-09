import { and, asc, desc, eq, gte } from 'drizzle-orm';

import { emitTableChanges, getDb, runInTransaction } from '@/core/db';
import { todayIso, type IsoDate } from '@/core/utils';
import {
  isWorkingSet,
  summarizeExerciseHistory,
  type ExercisePreview,
  type ExerciseSetRow,
  type HistorySetRow,
  type LoadType,
  type UnilateralCounting,
} from '@/domain/fitness';
import type { Workout, WorkoutExercise } from '@/domain/models';

import { newId } from '../id';
import { exercises, sets, workoutDrafts, workoutExercises, workouts } from '../schema/tables';

/** The single-row id for the crash-safe draft (DATABASE §3.4: `id = 1`). */
const DRAFT_ID = 1;

export interface SetEditPatch {
  readonly weightKg?: number;
  readonly reps?: number;
  readonly warmup?: boolean;
}

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
  /** Provenance: the template this session started from, or null (DATABASE §3.4). */
  readonly templateId?: string | null;
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
        templateId: input.templateId ?? null,
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

      // Finish is a single durable write that also clears the draft in the SAME
      // transaction (ARCHITECTURE §7.1.2): the saved workout and the vanished
      // draft commit together, so a crash can never leave both — or neither.
      await db.delete(workoutDrafts).where(eq(workoutDrafts.id, DRAFT_ID));
    });

    emitTableChanges('workouts');
    return workoutId;
  },

  /**
   * Writes the crash-safe checkpoint (ARCHITECTURE §7.1). Upsert of the single
   * draft row; high-frequency, so it does NOT touch the change-bus. The payload
   * is an opaque JSON string owned by the workouts feature — the repository does
   * not interpret it (keeping the session shape out of the data layer).
   */
  async checkpointDraft(payload: string): Promise<void> {
    await getDb()
      .insert(workoutDrafts)
      .values({ id: DRAFT_ID, payload, updatedAt: Date.now() })
      .onConflictDoUpdate({
        target: workoutDrafts.id,
        set: { payload, updatedAt: Date.now() },
      });
  },

  /** Loads the raw draft payload for recovery, or null when none exists. */
  async loadDraft(): Promise<string | null> {
    const rows = await getDb()
      .select({ payload: workoutDrafts.payload })
      .from(workoutDrafts)
      .where(eq(workoutDrafts.id, DRAFT_ID));
    return rows[0]?.payload ?? null;
  },

  /** Removes the draft (on explicit discard, or after recovering an invalid one). */
  async discardDraft(): Promise<void> {
    await getDb().delete(workoutDrafts).where(eq(workoutDrafts.id, DRAFT_ID));
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

  /**
   * Dates of **countable** workouts (≥ 1 working set, §3.8) since a date — one entry
   * per countable workout (two on a day → two entries). SQL only fetches the rows;
   * the domain `isWorkingSet` (which needs load type, not just the warm-up flag)
   * decides countability, so no domain semantics leak into SQL (ANALYTICS rule 9).
   * Feeds the dashboard streak + weekly-consistency (§3.8).
   */
  async getCountableWorkoutDatesSince(sinceIso: IsoDate): Promise<IsoDate[]> {
    const rows = await getDb()
      .select({
        workoutId: workouts.id,
        date: workouts.date,
        loadType: exercises.loadType,
        weightKg: sets.weightKg,
        reps: sets.reps,
        isWarmup: sets.isWarmup,
      })
      .from(sets)
      .innerJoin(workoutExercises, eq(sets.workoutExerciseId, workoutExercises.id))
      .innerJoin(exercises, eq(workoutExercises.exerciseId, exercises.id))
      .innerJoin(workouts, eq(workoutExercises.workoutId, workouts.id))
      .where(gte(workouts.date, sinceIso));

    const countable = new Map<string, IsoDate>();
    for (const r of rows) {
      if (countable.has(r.workoutId)) continue;
      if (
        isWorkingSet(
          { weightKg: r.weightKg, reps: r.reps, warmup: r.isWarmup === 1 },
          r.loadType as LoadType,
        )
      ) {
        countable.set(r.workoutId, r.date);
      }
    }
    return [...countable.values()];
  },

  /** Deletes a workout and its cascade tree. */
  async remove(id: string): Promise<void> {
    await getDb()
      .delete(workouts)
      .where(and(eq(workouts.id, id)));
    emitTableChanges('workouts');
  },

  /** Most recent workouts (full trees), newest first — powers the history list. */
  async listRecent(limit = 20): Promise<Workout[]> {
    const idRows = await getDb()
      .select({ id: workouts.id })
      .from(workouts)
      .orderBy(desc(workouts.date), desc(workouts.createdAt))
      .limit(limit);
    const result: Workout[] = [];
    for (const { id } of idRows) {
      const workout = await this.getById(id);
      if (workout) result.push(workout);
    }
    return result;
  },

  /** All logged sets of an exercise across history, newest workout first (for the preview). */
  async getExerciseHistory(exerciseId: string): Promise<HistorySetRow[]> {
    const rows = await getDb()
      .select({
        workoutId: workouts.id,
        date: workouts.date,
        workoutOrder: workouts.createdAt,
        position: sets.position,
        weightKg: sets.weightKg,
        reps: sets.reps,
        isWarmup: sets.isWarmup,
      })
      .from(sets)
      .innerJoin(workoutExercises, eq(sets.workoutExerciseId, workoutExercises.id))
      .innerJoin(workouts, eq(workoutExercises.workoutId, workouts.id))
      .where(eq(workoutExercises.exerciseId, exerciseId))
      .orderBy(desc(workouts.createdAt), asc(sets.position));
    return rows.map((r) => ({
      workoutId: r.workoutId,
      date: r.date,
      workoutOrder: r.workoutOrder,
      weightKg: r.weightKg,
      reps: r.reps,
      warmup: r.isWarmup === 1,
    }));
  },

  /** Last / Best / Best e1RM for an exercise (FITNESS_DOMAIN §3, computed in the domain). */
  async getExercisePreview(exerciseId: string, loadType: LoadType): Promise<ExercisePreview> {
    const rows = await this.getExerciseHistory(exerciseId);
    return summarizeExerciseHistory(rows, loadType);
  },

  /**
   * Full working-history rows for an exercise (with per-entry unilateral counting)
   * plus its name/load type — the raw input the domain turns into PRs (§3.7) and
   * the Exercise Report (§5.5). Returns null when the exercise does not exist.
   */
  async getExerciseSetHistory(
    exerciseId: string,
  ): Promise<{ name: string; loadType: LoadType; rows: ExerciseSetRow[] } | null> {
    const db = getDb();
    const exerciseRows = await db
      .select({ name: exercises.name, loadType: exercises.loadType })
      .from(exercises)
      .where(eq(exercises.id, exerciseId));
    const exercise = exerciseRows[0];
    if (!exercise) return null;

    const rows = await db
      .select({
        workoutId: workouts.id,
        date: workouts.date,
        workoutOrder: workouts.createdAt,
        counting: workoutExercises.unilateralCounting,
        weightKg: sets.weightKg,
        reps: sets.reps,
        isWarmup: sets.isWarmup,
      })
      .from(sets)
      .innerJoin(workoutExercises, eq(sets.workoutExerciseId, workoutExercises.id))
      .innerJoin(workouts, eq(workoutExercises.workoutId, workouts.id))
      .where(eq(workoutExercises.exerciseId, exerciseId))
      .orderBy(desc(workouts.createdAt), asc(sets.position));

    return {
      name: exercise.name,
      loadType: exercise.loadType as LoadType,
      rows: rows.map((r) => ({
        workoutId: r.workoutId,
        date: r.date,
        workoutOrder: r.workoutOrder,
        weightKg: r.weightKg,
        reps: r.reps,
        warmup: r.isWarmup === 1,
        counting: r.counting as UnilateralCounting,
      })),
    };
  },

  /** Edits a saved set; derived views recompute via the change-bus (ARCHITECTURE rule 8). */
  async updateSet(setId: string, patch: SetEditPatch): Promise<void> {
    await getDb()
      .update(sets)
      .set({
        ...(patch.weightKg !== undefined && { weightKg: patch.weightKg }),
        ...(patch.reps !== undefined && { reps: patch.reps }),
        ...(patch.warmup !== undefined && { isWarmup: patch.warmup ? 1 : 0 }),
      })
      .where(eq(sets.id, setId));
    emitTableChanges('workouts');
  },

  /** Deletes a single saved set. */
  async deleteSet(setId: string): Promise<void> {
    await getDb().delete(sets).where(eq(sets.id, setId));
    emitTableChanges('workouts');
  },
} as const;
