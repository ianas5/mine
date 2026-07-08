import Database from 'better-sqlite3';
import { drizzle, type BetterSQLite3Database } from 'drizzle-orm/better-sqlite3';
import { migrate } from 'drizzle-orm/better-sqlite3/migrator';

/**
 * Node-only test harness (DEVELOPMENT_WORKFLOW §4.2): a real in-memory SQLite
 * database running the exact same generated migrations as the device runtime.
 * Use with `setDbForTesting` from @/core/db in `@jest-environment node` suites.
 */
export interface TestDb {
  readonly db: BetterSQLite3Database;
  readonly sqlite: Database.Database;
  readonly close: () => void;
}

export function createTestDb(): TestDb {
  const sqlite = new Database(':memory:');
  sqlite.pragma('foreign_keys = ON');
  const db = drizzle(sqlite);
  migrate(db, { migrationsFolder: 'src/data/schema/migrations' });
  return { db, sqlite, close: () => sqlite.close() };
}
