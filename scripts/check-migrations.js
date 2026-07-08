/**
 * CI migration check (DEVELOPMENT_WORKFLOW §3): all committed migrations apply
 * cleanly to a fresh database. Fixture-upgrade tests join per-migration PRs.
 */
const Database = require('better-sqlite3');
const { drizzle } = require('drizzle-orm/better-sqlite3');
const { migrate } = require('drizzle-orm/better-sqlite3/migrator');

const sqlite = new Database(':memory:');
sqlite.pragma('foreign_keys = ON');
migrate(drizzle(sqlite), { migrationsFolder: 'src/data/schema/migrations' });

const tables = sqlite
  .prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE '__drizzle%'")
  .all()
  .map((row) => row.name)
  .sort();

console.log(`migrations OK — tables: ${tables.join(', ')}`);
sqlite.close();
