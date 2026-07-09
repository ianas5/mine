import {
  BACKUP_APP,
  BACKUP_FORMAT,
  backupDataSchema,
  backupHeaderSchema,
  type BackupEnvelope,
} from './backupSchema';
import { ImportError } from './importError';
import { CURRENT_SCHEMA_VERSION } from './schemaVersion';
import { upgradeData } from './upgraders';

/**
 * The full import validation gate (DATABASE §6.2–6.3), pure and side-effect-free so
 * it can run — and be exhaustively tested — before anything is written. Order:
 *
 *   1. JSON.parse → `unreadable`/`invalid-data` on malformed text.
 *   2. Header parse (lenient version fields) + app/format check.
 *   3. Version reconcile: refuse newer, upgrade older, accept current.
 *   4. Full Zod validation of the (upgraded) `data` against the current shape.
 *
 * Any failure throws an `ImportError` and, by construction, touches no data.
 */
export function parseAndUpgradeEnvelope(dataJson: string): BackupEnvelope {
  let raw: unknown;
  try {
    raw = JSON.parse(dataJson);
  } catch {
    throw new ImportError('invalid-data', 'data.json is not valid JSON');
  }

  const header = backupHeaderSchema.safeParse(raw);
  if (!header.success) {
    throw new ImportError('invalid-data', 'backup header is missing or malformed');
  }

  const { app, format, schemaVersion, exportedAt, data } = header.data;
  if (app !== BACKUP_APP) {
    throw new ImportError('invalid-data', `not a ${BACKUP_APP} backup`);
  }
  if (format !== BACKUP_FORMAT) {
    throw new ImportError('unsupported-format', `unsupported archive format ${format}`);
  }

  // Version reconcile before shape validation (upgraders target the current shape).
  const upgraded = upgradeData((data ?? {}) as Record<string, unknown>, schemaVersion);

  const parsed = backupDataSchema.safeParse(upgraded);
  if (!parsed.success) {
    throw new ImportError('invalid-data', 'backup data failed validation');
  }

  return {
    app: BACKUP_APP,
    format: BACKUP_FORMAT,
    schemaVersion: CURRENT_SCHEMA_VERSION,
    exportedAt,
    data: parsed.data,
  };
}
