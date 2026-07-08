import { eq } from 'drizzle-orm';

import { emitTableChanges, getDb } from '@/core/db';
import type { Settings } from '@/domain/models';

import { rowToSettings } from '../mappers/settingsMapper';
import { settings } from '../schema/tables';

const SINGLETON_ID = 1;

/**
 * Single-row domain configuration (DATABASE §3.1/§7): get creates the row if
 * missing; update is upsert-only and emits the change-bus as its last step.
 */
export const settingsRepository = {
  async get(): Promise<Settings> {
    const db = getDb();
    const rows = await db.select().from(settings).where(eq(settings.id, SINGLETON_ID));
    const existing = rows[0];
    if (existing) {
      return rowToSettings(existing);
    }
    const now = Date.now();
    const inserted = await db
      .insert(settings)
      .values({ id: SINGLETON_ID, createdAt: now, updatedAt: now })
      .returning();
    const created = inserted[0];
    if (!created) {
      throw new Error('settings row could not be created');
    }
    return rowToSettings(created);
  },

  async update(patch: Partial<Settings>): Promise<void> {
    await this.get(); // guarantees the row exists
    const db = getDb();
    await db
      .update(settings)
      .set({ ...patch, updatedAt: Date.now() })
      .where(eq(settings.id, SINGLETON_ID));
    emitTableChanges('settings');
  },
} as const;
