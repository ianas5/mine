/** @jest-environment node */
import journal from '../schema/migrations/meta/_journal.json';
import { CURRENT_SCHEMA_VERSION } from './schemaVersion';

describe('CURRENT_SCHEMA_VERSION', () => {
  it('equals the number of applied migrations (no drift)', () => {
    // Bumping a migration without bumping this constant would silently mislabel
    // every export's schemaVersion — this guard fails the build if they diverge.
    expect(CURRENT_SCHEMA_VERSION).toBe(journal.entries.length);
  });
});
