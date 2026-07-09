/**
 * The DB migration version stamped into every backup (`schemaVersion`, DATABASE §6).
 * It equals the number of applied Drizzle migrations (`meta/_journal.json` entries):
 * a monotonic integer that increases by one with every schema-changing migration.
 *
 * On import it decides upgrade-vs-refuse: an archive at this version needs no
 * upgrade; an older one runs the data-shape upgraders up to it; a newer one is
 * refused ("update the app first"). It is deliberately a hand-maintained constant
 * — bumped in the same commit as a new migration — and guarded against drift by
 * `schemaVersion.test.ts`, which asserts it matches the migration journal length.
 */
export const CURRENT_SCHEMA_VERSION = 11;
