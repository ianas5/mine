import { ImportError } from './importError';
import { CURRENT_SCHEMA_VERSION } from './schemaVersion';

/**
 * Data-shape upgraders (DATABASE §6.3): pure functions that migrate a decoded
 * `data` block from one `schemaVersion` to the next, mirroring the semantics of the
 * SQL migration that bumped the version. They run in sequence from the archive's
 * version up to `CURRENT_SCHEMA_VERSION` **before** the full Zod validation, so an
 * old backup is validated against — and inserted as — the current shape.
 *
 * The map is keyed by the version being upgraded FROM; each entry produces the shape
 * at version `key + 1`. It is empty at v1 (only the current schema has ever shipped);
 * the first entry arrives with the `phases` table in Phase 19. A gap in the chain
 * (an old version with no registered step) is refused, never guessed.
 */
type Upgrader = (data: Record<string, unknown>) => Record<string, unknown>;

const UPGRADERS: Readonly<Record<number, Upgrader>> = {
  // v9 → v10 (migration 0009): settings gained `target_weight_kg`. An older backup's
  // settings object has no such key; default it to null (no goal set) so the current
  // Zod shape validates. Nested inside `data.settings`, which is a single object.
  9: (data) => {
    const settings = (data.settings ?? {}) as Record<string, unknown>;
    return { ...data, settings: { targetWeightKg: null, ...settings } };
  },
  // v10 → v11 (migration 0010): the `phases` table was added (Phase 19). An older
  // backup has no `phases` array; default it to empty (no declared phases) so the
  // current Zod shape validates. Phases are additive — no existing data changes.
  10: (data) => ({ phases: [], ...data }),
};

/**
 * Brings a decoded `data` block from `fromVersion` up to the current schema, or
 * throws `ImportError` if the archive is newer than this app or an intermediate
 * upgrader is missing. Returns the (still untyped) upgraded block for Zod to validate.
 */
export function upgradeData(
  data: Record<string, unknown>,
  fromVersion: number,
): Record<string, unknown> {
  if (fromVersion > CURRENT_SCHEMA_VERSION) {
    throw new ImportError('schema-too-new');
  }

  let current = data;
  for (let version = fromVersion; version < CURRENT_SCHEMA_VERSION; version += 1) {
    const step = UPGRADERS[version];
    if (!step) {
      throw new ImportError('unsupported-schema');
    }
    current = step(current);
  }
  return current;
}
