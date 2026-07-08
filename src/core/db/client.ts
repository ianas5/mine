import { drizzle, type ExpoSQLiteDatabase } from 'drizzle-orm/expo-sqlite';
import { openDatabaseSync } from 'expo-sqlite';

import type { BetterSQLite3Database } from 'drizzle-orm/better-sqlite3';

/**
 * The app database handle. Native runtime uses the expo-sqlite driver; the
 * test harness injects a better-sqlite3-backed instance with the identical
 * schema/migrations (DEVELOPMENT_WORKFLOW §4.2). Both satisfy the same
 * query-builder surface used by repositories.
 */
export type AppDb = ExpoSQLiteDatabase | BetterSQLite3Database;

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

/** Test-only injection point for the better-sqlite3 harness. */
export function setDbForTesting(testDb: AppDb): void {
  db = testDb;
}
