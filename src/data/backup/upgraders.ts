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
  // 9: (data) => ({ ...data, phases: [] }),   // example — added in Phase 19
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
