import { asc, desc, eq, isNull } from 'drizzle-orm';

import { emitTableChanges, getDb, runInTransaction } from '@/core/db';
import { type IsoDate } from '@/core/utils';
import type { Phase, PhaseType } from '@/domain/models';

import { newId } from '../id';
import { phases } from '../schema/tables';

/** Far-future sentinel standing in for an ongoing phase's open end (ISO dates sort
 * lexicographically, so this compares as "after" every real date). */
const OPEN_END = '9999-12-31';

export type PhaseValidationReason = 'overlap' | 'end-before-start';

/** Thrown when a create/update would break the no-overlap or ordering invariant
 * (DATABASE §3.7). The UI catches it to guide the fix (e.g. end the current phase). */
export class PhaseValidationError extends Error {
  constructor(readonly reason: PhaseValidationReason) {
    super(reason);
    this.name = 'PhaseValidationError';
  }
}

export interface PhaseInput {
  readonly name: string;
  readonly type: PhaseType;
  readonly startDate: IsoDate;
  readonly endDate: IsoDate | null;
  readonly notes: string | null;
}

function toPhase(row: typeof phases.$inferSelect): Phase {
  return {
    id: row.id,
    name: row.name,
    type: row.type as PhaseType,
    startDate: row.startDate,
    endDate: row.endDate,
    notes: row.notes,
    createdAt: row.createdAt,
    updatedAt: row.updatedAt,
  };
}

/** Do two (possibly open-ended) date intervals share at least one day? */
function overlaps(
  aStart: IsoDate,
  aEnd: IsoDate | null,
  bStart: IsoDate,
  bEnd: IsoDate | null,
): boolean {
  return aStart <= (bEnd ?? OPEN_END) && bStart <= (aEnd ?? OPEN_END);
}

/** Reads every phase except `exceptId` and throws if `range` overlaps any of them.
 * Also collapses the single-ongoing rule: a second open-ended phase always overlaps
 * the first (both extend to `OPEN_END`). Must run inside the write transaction. */
async function assertNoOverlap(
  range: { startDate: IsoDate; endDate: IsoDate | null },
  exceptId: string | null,
): Promise<void> {
  if (range.endDate !== null && range.endDate < range.startDate) {
    throw new PhaseValidationError('end-before-start');
  }
  const others = await getDb().select().from(phases);
  for (const other of others) {
    if (other.id === exceptId) continue;
    if (overlaps(range.startDate, range.endDate, other.startDate, other.endDate)) {
      throw new PhaseValidationError('overlap');
    }
  }
}

export const phaseRepository = {
  /** All phases, most recent first (ongoing sorts first via its late start date). */
  async listPhases(): Promise<Phase[]> {
    const rows = await getDb().select().from(phases).orderBy(desc(phases.startDate));
    return rows.map(toPhase);
  },

  async getPhase(id: string): Promise<Phase | null> {
    const rows = await getDb().select().from(phases).where(eq(phases.id, id));
    const row = rows[0];
    return row ? toPhase(row) : null;
  },

  /** The single ongoing phase (`end_date IS NULL`), or null. */
  async getOngoingPhase(): Promise<Phase | null> {
    const rows = await getDb().select().from(phases).where(isNull(phases.endDate));
    const row = rows[0];
    return row ? toPhase(row) : null;
  },

  /** The phase that was active on `date` — the one whose range covers it. Supports the
   * rule that historical results are read against the phase active at that time; returns
   * null for a date in no declared phase. (No-overlap guarantees at most one match.) */
  async getPhaseForDate(date: IsoDate): Promise<Phase | null> {
    const rows = await getDb().select().from(phases).orderBy(asc(phases.startDate));
    const match = rows.find(
      (r) => r.startDate <= date && (r.endDate === null || date <= r.endDate),
    );
    return match ? toPhase(match) : null;
  },

  /**
   * Declares a phase, enforcing no-overlap + single-ongoing (DATABASE §3.7) inside a
   * transaction. Throws `PhaseValidationError` if the range collides with an existing
   * phase — the caller ends the current phase first (end-yesterday UX), then retries.
   */
  async createPhase(input: PhaseInput): Promise<string> {
    const id = newId('phase');
    const now = Date.now();
    await runInTransaction(async () => {
      await assertNoOverlap({ startDate: input.startDate, endDate: input.endDate }, null);
      await getDb()
        .insert(phases)
        .values({
          id,
          name: input.name.trim() || 'Phase',
          type: input.type,
          startDate: input.startDate,
          endDate: input.endDate,
          notes: input.notes,
          createdAt: now,
          updatedAt: now,
        });
    });
    emitTableChanges('phases');
    return id;
  },

  /** Edits a phase's fields, re-checking the invariants against the other phases. */
  async updatePhase(
    id: string,
    patch: Partial<Pick<PhaseInput, 'name' | 'type' | 'startDate' | 'endDate' | 'notes'>>,
  ): Promise<void> {
    await runInTransaction(async () => {
      const db = getDb();
      const rows = await db.select().from(phases).where(eq(phases.id, id));
      const existing = rows[0];
      if (!existing) return;
      const startDate = patch.startDate ?? existing.startDate;
      const endDate = patch.endDate !== undefined ? patch.endDate : existing.endDate;
      await assertNoOverlap({ startDate, endDate }, id);
      await db
        .update(phases)
        .set({
          ...(patch.name !== undefined && { name: patch.name.trim() || 'Phase' }),
          ...(patch.type !== undefined && { type: patch.type }),
          ...(patch.startDate !== undefined && { startDate }),
          ...(patch.endDate !== undefined && { endDate }),
          ...(patch.notes !== undefined && { notes: patch.notes }),
          updatedAt: Date.now(),
        })
        .where(eq(phases.id, id));
    });
    emitTableChanges('phases');
  },

  /** Closes an ongoing phase on `endDate` (the UI defaults this to yesterday). Ending
   * only shrinks the interval, so it can never introduce an overlap — but the
   * ordering (`end ≥ start`) is still validated. */
  async endPhase(id: string, endDate: IsoDate): Promise<void> {
    await this.updatePhase(id, { endDate });
  },

  async deletePhase(id: string): Promise<void> {
    await getDb().delete(phases).where(eq(phases.id, id));
    emitTableChanges('phases');
  },
} as const;
