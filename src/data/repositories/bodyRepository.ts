import { desc, eq, isNotNull } from 'drizzle-orm';

import { emitTableChanges, getDb } from '@/core/db';
import type { IsoDate } from '@/core/utils';
import { BODY_FIELDS, type BodyField, type BodySnapshot } from '@/domain/body';

import { bodySnapshots } from '../schema/tables';

/** A save patch: only fields to *set* (numbers). It cannot express a clear — that is
 * the deliberate `clearField` action, so an omitted field can never erase a value. */
export type MeasurementPatch = Partial<Record<BodyField, number>>;

function rowToSnapshot(row: typeof bodySnapshots.$inferSelect): BodySnapshot {
  const values = {} as Record<BodyField, number | null>;
  for (const field of BODY_FIELDS) values[field] = row[field];
  return { date: row.date, ...values };
}

export const bodyRepository = {
  /** One date's snapshot, or null. */
  async getSnapshot(date: IsoDate): Promise<BodySnapshot | null> {
    const rows = await getDb().select().from(bodySnapshots).where(eq(bodySnapshots.date, date));
    const row = rows[0];
    return row ? rowToSnapshot(row) : null;
  },

  /** All snapshots, newest first (feeds current-state, placeholders, frequency). */
  async listSnapshots(): Promise<BodySnapshot[]> {
    const rows = await getDb().select().from(bodySnapshots).orderBy(desc(bodySnapshots.date));
    return rows.map(rowToSnapshot);
  },

  /** Weigh-ins (date + weight) newest first — the input to the delta'd weight log. */
  async getWeightLog(): Promise<{ date: IsoDate; weightKg: number }[]> {
    const rows = await getDb()
      .select({ date: bodySnapshots.date, weightKg: bodySnapshots.weightKg })
      .from(bodySnapshots)
      .where(isNotNull(bodySnapshots.weightKg))
      .orderBy(desc(bodySnapshots.date));
    return rows.flatMap((r) =>
      r.weightKg === null ? [] : [{ date: r.date, weightKg: r.weightKg }],
    );
  },

  /**
   * Merge-upsert a snapshot (DATABASE §3.6, FITNESS_DOMAIN §5.1): writes only the
   * fields present in `patch`. On an existing date, omitted fields keep their stored
   * values — an omission NEVER clears. On a new date, absent fields default to NULL.
   */
  async saveSnapshot(date: IsoDate, patch: MeasurementPatch): Promise<void> {
    const changes: Partial<Record<BodyField, number>> = {};
    for (const field of BODY_FIELDS) {
      const value = patch[field];
      if (value !== undefined) changes[field] = value;
    }

    const now = Date.now();
    await getDb()
      .insert(bodySnapshots)
      .values({ date, ...changes, createdAt: now, updatedAt: now })
      .onConflictDoUpdate({
        target: bodySnapshots.date,
        set: { ...changes, updatedAt: now },
      });
    emitTableChanges('body');
  },

  /**
   * Explicitly clears one field on a date (the deliberate action, DATABASE §3.6) —
   * the only way a stored value is nulled. No-op if the snapshot does not exist.
   */
  async clearField(date: IsoDate, field: BodyField): Promise<void> {
    await getDb()
      .update(bodySnapshots)
      .set({ [field]: null, updatedAt: Date.now() })
      .where(eq(bodySnapshots.date, date));
    emitTableChanges('body');
  },
} as const;
