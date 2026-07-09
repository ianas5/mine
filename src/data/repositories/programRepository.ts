import { and, asc, desc, eq, gte, isNull } from 'drizzle-orm';

import { emitTableChanges, getDb, runInTransaction } from '@/core/db';
import { isoWeekday } from '@/core/utils';
import type { LoadType, RecentTemplateUse } from '@/domain/fitness';
import type { Program, Template, TemplateExercise } from '@/domain/models';

import { newId } from '../id';
import { exercises, programs, templateExercises, templates, workouts } from '../schema/tables';

export interface TemplateExerciseInput {
  readonly exerciseId: string;
  readonly targetSets: number | null;
  readonly targetRepMin: number | null;
  readonly targetRepMax: number | null;
  readonly targetRpe: number | null;
  readonly restSeconds: number | null;
  readonly notes: string | null;
}

export interface TemplateInput {
  readonly name: string;
  readonly weekday: number | null;
  readonly notes: string | null;
  readonly exercises: readonly TemplateExerciseInput[];
}

const num = (value: number | null): number | null => (value === null ? null : value);

async function loadTemplateExercises(templateId: string): Promise<TemplateExercise[]> {
  const rows = await getDb()
    .select({
      id: templateExercises.id,
      exerciseId: templateExercises.exerciseId,
      name: exercises.name,
      loadType: exercises.loadType,
      defaultUnilateral: exercises.defaultUnilateral,
      position: templateExercises.position,
      targetSets: templateExercises.targetSets,
      targetRepMin: templateExercises.targetRepMin,
      targetRepMax: templateExercises.targetRepMax,
      targetRpe: templateExercises.targetRpe,
      restSeconds: templateExercises.restSeconds,
      notes: templateExercises.notes,
    })
    .from(templateExercises)
    .innerJoin(exercises, eq(templateExercises.exerciseId, exercises.id))
    .where(eq(templateExercises.templateId, templateId))
    .orderBy(asc(templateExercises.position));

  return rows.map((r) => ({
    id: r.id,
    exerciseId: r.exerciseId,
    name: r.name,
    loadType: r.loadType as LoadType,
    defaultUnilateral: r.defaultUnilateral === 1,
    position: r.position,
    target: {
      sets: r.targetSets,
      repMin: r.targetRepMin,
      repMax: r.targetRepMax,
      rpe: r.targetRpe,
      restSeconds: r.restSeconds,
    },
    notes: r.notes,
  }));
}

function toTemplate(
  row: typeof templates.$inferSelect,
  exerciseList: TemplateExercise[],
): Template {
  return {
    id: row.id,
    programId: row.programId,
    name: row.name,
    position: row.position,
    weekday: row.weekday,
    notes: row.notes,
    exercises: exerciseList,
  };
}

async function loadTemplatesFor(programId: string | null): Promise<Template[]> {
  const rows = await getDb()
    .select()
    .from(templates)
    .where(
      programId === null
        ? and(isNull(templates.programId), eq(templates.isArchived, 0))
        : and(eq(templates.programId, programId), eq(templates.isArchived, 0)),
    )
    .orderBy(asc(templates.position));

  const built: Template[] = [];
  for (const row of rows) built.push(toTemplate(row, await loadTemplateExercises(row.id)));
  return built;
}

function toProgram(row: typeof programs.$inferSelect, templateList: Template[]): Program {
  return {
    id: row.id,
    name: row.name,
    notes: row.notes,
    isActive: row.isActive === 1,
    isArchived: row.isArchived === 1,
    createdAt: row.createdAt,
    updatedAt: row.updatedAt,
    templates: templateList,
  };
}

async function insertTemplateExercises(
  templateId: string,
  list: readonly TemplateExerciseInput[],
): Promise<void> {
  const db = getDb();
  for (const [position, te] of list.entries()) {
    await db.insert(templateExercises).values({
      id: newId('te'),
      templateId,
      exerciseId: te.exerciseId,
      position,
      targetSets: num(te.targetSets),
      targetRepMin: num(te.targetRepMin),
      targetRepMax: num(te.targetRepMax),
      targetRpe: num(te.targetRpe),
      restSeconds: num(te.restSeconds),
      notes: te.notes,
    });
  }
}

export const programRepository = {
  /** Non-archived programs (each with its templates), newest first. */
  async listPrograms(): Promise<Program[]> {
    const rows = await getDb()
      .select()
      .from(programs)
      .where(eq(programs.isArchived, 0))
      .orderBy(asc(programs.name));
    const built: Program[] = [];
    for (const row of rows) built.push(toProgram(row, await loadTemplatesFor(row.id)));
    return built;
  },

  /** The single active program (DATABASE §3.3), or null. */
  async getActiveProgram(): Promise<Program | null> {
    const rows = await getDb()
      .select()
      .from(programs)
      .where(and(eq(programs.isActive, 1), eq(programs.isArchived, 0)));
    const row = rows[0];
    if (!row) return null;
    return toProgram(row, await loadTemplatesFor(row.id));
  },

  async getProgram(id: string): Promise<Program | null> {
    const rows = await getDb().select().from(programs).where(eq(programs.id, id));
    const row = rows[0];
    if (!row) return null;
    return toProgram(row, await loadTemplatesFor(row.id));
  },

  async createProgram(input: { name: string; notes: string | null }): Promise<string> {
    const id = newId('pg');
    const now = Date.now();
    await getDb()
      .insert(programs)
      .values({
        id,
        name: input.name.trim() || 'Program',
        notes: input.notes,
        isActive: 0,
        isArchived: 0,
        createdAt: now,
        updatedAt: now,
      });
    emitTableChanges('programs');
    return id;
  },

  async updateProgram(id: string, patch: { name?: string; notes?: string | null }): Promise<void> {
    await getDb()
      .update(programs)
      .set({
        ...(patch.name !== undefined && { name: patch.name.trim() || 'Program' }),
        ...(patch.notes !== undefined && { notes: patch.notes }),
        updatedAt: Date.now(),
      })
      .where(eq(programs.id, id));
    emitTableChanges('programs');
  },

  /**
   * Makes one program active, enforcing the single-active invariant (DATABASE §3.3)
   * in a transaction: clear every program's flag, then set this one.
   */
  async setActive(id: string): Promise<void> {
    await runInTransaction(async () => {
      const db = getDb();
      await db.update(programs).set({ isActive: 0, updatedAt: Date.now() });
      await db
        .update(programs)
        .set({ isActive: 1, updatedAt: Date.now() })
        .where(eq(programs.id, id));
    });
    emitTableChanges('programs');
  },

  /** Clears the active program (no program active). */
  async clearActive(): Promise<void> {
    await getDb().update(programs).set({ isActive: 0, updatedAt: Date.now() });
    emitTableChanges('programs');
  },

  /** Archives a program (kept for provenance; its templates are hidden with it). */
  async archiveProgram(id: string): Promise<void> {
    await getDb()
      .update(programs)
      .set({ isArchived: 1, isActive: 0, updatedAt: Date.now() })
      .where(eq(programs.id, id));
    emitTableChanges('programs');
  },

  /** Hard-deletes a program and its templates (cascade). Past workouts keep their data. */
  async deleteProgram(id: string): Promise<void> {
    await getDb().delete(programs).where(eq(programs.id, id));
    emitTableChanges('programs');
  },

  async getTemplate(id: string): Promise<Template | null> {
    const rows = await getDb().select().from(templates).where(eq(templates.id, id));
    const row = rows[0];
    if (!row) return null;
    return toTemplate(row, await loadTemplateExercises(id));
  },

  /** Creates a template (standalone when programId is null) with its exercises. */
  async createTemplate(programId: string | null, input: TemplateInput): Promise<string> {
    const id = newId('tpl');
    const now = Date.now();
    await runInTransaction(async () => {
      const db = getDb();
      const existing = await db
        .select({ id: templates.id })
        .from(templates)
        .where(
          programId === null ? isNull(templates.programId) : eq(templates.programId, programId),
        );
      await db.insert(templates).values({
        id,
        programId,
        name: input.name.trim() || 'Session',
        position: existing.length,
        weekday: input.weekday,
        notes: input.notes,
        isArchived: 0,
        createdAt: now,
        updatedAt: now,
      });
      await insertTemplateExercises(id, input.exercises);
    });
    emitTableChanges('programs');
    return id;
  },

  /**
   * Replaces a template's fields and its exercise list (DATABASE §3.3). This edits
   * the *plan only* — template_exercises rows are rebuilt; no workout is touched.
   */
  async updateTemplate(id: string, input: TemplateInput): Promise<void> {
    await runInTransaction(async () => {
      const db = getDb();
      await db
        .update(templates)
        .set({
          name: input.name.trim() || 'Session',
          weekday: input.weekday,
          notes: input.notes,
          updatedAt: Date.now(),
        })
        .where(eq(templates.id, id));
      await db.delete(templateExercises).where(eq(templateExercises.templateId, id));
      await insertTemplateExercises(id, input.exercises);
    });
    emitTableChanges('programs');
  },

  /** Deletes a template (cascade to its exercises). Past workouts keep their history. */
  async deleteTemplate(id: string): Promise<void> {
    await getDb().delete(templates).where(eq(templates.id, id));
    emitTableChanges('programs');
  },

  /**
   * Template-started workouts within the lookback window (for the weekday-mode
   * suggestion, UI_UX §5.2). Only rows whose template still exists (template_id not
   * SET NULL by a delete) are returned, newest first.
   */
  async getRecentTemplateUses(sinceIso: string): Promise<RecentTemplateUse[]> {
    const rows = await getDb()
      .select({ templateId: workouts.templateId, date: workouts.date })
      .from(workouts)
      .innerJoin(templates, eq(workouts.templateId, templates.id))
      .where(and(gte(workouts.date, sinceIso), eq(templates.isArchived, 0)))
      .orderBy(desc(workouts.date), desc(workouts.createdAt));

    return rows
      .filter((r): r is { templateId: string; date: string } => r.templateId !== null)
      .map((r) => ({ weekday: isoWeekday(r.date), templateId: r.templateId }));
  },
} as const;
