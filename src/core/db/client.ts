import { sql } from 'drizzle-orm';
import { drizzle, type ExpoSQLiteDatabase } from 'drizzle-orm/expo-sqlite';
import { openDatabaseSync } from 'expo-sqlite';

import type { BetterSQLite3Database } from 'drizzle-orm/better-sqlite3';

/**
 * The app database handle. Typed as the production expo-sqlite driver so Drizzle's
 * query-builder generics infer cleanly (a union of the two drivers collapses
 * that inference). The better-sqlite3 test harness injects a structurally
 * compatible instance via `setDbForTesting` (DEVELOPMENT_WORKFLOW §4.2) — both
 * are Drizzle `BaseSQLiteDatabase` subclasses exposing the same builder surface;
 * `await` transparently handles the sync (test) vs async (device) result.
 */
export type AppDb = ExpoSQLiteDatabase;

let db: AppDb | null = null;

/** Opens the on-device database with DATABASE §1 PRAGMAs. Idempotent. */
export function initDb(): AppDb {
  if (db !== null) {
    return db;
  }
  const sqlite = openDatabaseSync('fitness.db');
  sqlite.execSync('PRAGMA journal_mode = WAL;');
  sqlite.execSync('PRAGMA foreign_keys = ON;');
  sqlite.execSync('PRAGMA busy_timeout = 5000;');
  db = drizzle(sqlite);
  return db;
}

/** The shared connection (ARCHITECTURE §7 read/write paths). Init must have run. */
export function getDb(): AppDb {
  if (db === null) {
    throw new Error('Database not initialized — the DB-ready gate must run before any query.');
  }
  return db;
}

/** Test-only injection point for the better-sqlite3 harness (structurally compatible). */
export function setDbForTesting(testDb: BetterSQLite3Database): void {
  db = testDb as unknown as AppDb;
}

/**
 * Runs `fn` inside a single all-or-nothing transaction (ARCHITECTURE rule 11).
 * Manual BEGIN/COMMIT/ROLLBACK via `run` so it works identically over the
 * async expo-sqlite driver and the sync better-sqlite3 test driver.
 */
export async function runInTransaction<T>(fn: () => Promise<T>): Promise<T> {
  const active = getDb();
  await active.run(sql`BEGIN`);
  try {
    const result = await fn();
    await active.run(sql`COMMIT`);
    return result;
  } catch (cause) {
    try {
      await active.run(sql`ROLLBACK`);
    } catch {
      // ignore rollback failure; surface the original cause
    }
    throw cause;
  }
}
