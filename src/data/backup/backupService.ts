import { photoRepository } from '../repositories/photoRepository';
import { getArchiveStore } from './archiveStore';
import { BACKUP_APP, BACKUP_FORMAT, type BackupEnvelope } from './backupSchema';
import { collectBackupData } from './collect';
import { ImportError } from './importError';
import { parseAndUpgradeEnvelope } from './parseEnvelope';
import { replaceAllData } from './replace';
import { CURRENT_SCHEMA_VERSION } from './schemaVersion';

const isoDay = (iso: string): string => iso.slice(0, 10); // YYYY-MM-DD

/** Builds the current-DB backup envelope and packs it into an archive; returns its URI. */
async function packBackup(prefix: 'fitness-backup' | 'pre-import-safety'): Promise<string> {
  const data = await collectBackupData();
  const exportedAt = new Date().toISOString();
  const envelope: BackupEnvelope = {
    app: BACKUP_APP,
    format: BACKUP_FORMAT,
    schemaVersion: CURRENT_SCHEMA_VERSION,
    exportedAt,
    data,
  };
  return getArchiveStore().pack({
    dataJson: JSON.stringify(envelope),
    photoNames: data.progressPhotos.map((p) => p.fileName),
    fileName: `${prefix}-${isoDay(exportedAt)}.zip`,
  });
}

export interface ImportHooks {
  /**
   * Called only when the pre-import safety export fails. Must resolve to whether the
   * user explicitly confirmed continuing WITHOUT a safety backup. Silent continuation
   * is forbidden (DATABASE §6.7) — returning false aborts with nothing touched.
   */
  readonly onSafetyExportFailed: () => Promise<boolean>;
}

/**
 * The backup service (DATABASE §6) — a data-layer service, not a repository. Export
 * dumps every table + photos to a shareable zip. Import is all-or-nothing:
 * validate → version-reconcile → **attempted** safety export (failure ⇒ explicit
 * user confirmation) → single-transaction replace → photo reconcile. Existing data
 * is never touched unless validation passed and, if the safety export failed, the
 * user confirmed.
 */
export const backupService = {
  /** Builds a full backup archive and hands it to the OS share sheet. */
  async exportAndShare(): Promise<void> {
    const uri = await packBackup('fitness-backup');
    await getArchiveStore().share(uri);
  },

  /** Lets the user pick an archive to import; null if they cancel the picker. */
  async pickArchive(): Promise<string | null> {
    return getArchiveStore().pick();
  },

  /**
   * Imports an archive, replacing all data. Throws `ImportError` (nothing touched) on
   * any validation/version failure or a declined safety-export confirmation. On a
   * mid-replace fault the transaction rolls back — prior data stays intact.
   */
  async importArchive(uri: string, hooks: ImportHooks): Promise<void> {
    const store = getArchiveStore();
    try {
      let contents;
      try {
        contents = await store.open(uri);
      } catch {
        throw new ImportError('unreadable-archive', 'could not open the archive');
      }

      // Validate + reconcile version BEFORE any write and before the safety export, so
      // a garbage file fails fast without producing a spurious safety backup.
      const envelope = parseAndUpgradeEnvelope(contents.dataJson);

      // Attempt an automatic safety export of the CURRENT data before replacing it.
      let safetyOk = true;
      try {
        await packBackup('pre-import-safety');
      } catch {
        safetyOk = false;
      }
      if (!safetyOk) {
        const proceed = await hooks.onSafetyExportFailed();
        if (!proceed) {
          throw new ImportError('aborted-no-safety');
        }
      }

      // Point of no return: replace all rows atomically, then reconcile photo files.
      await replaceAllData(envelope.data);
      await store.commitPhotos(contents.photoNames);
      // Files without a row (from the replaced-away data) are removed; rows whose
      // file did not arrive are flagged missing on read (DATABASE §6.5).
      await photoRepository.sweepOrphans();
    } finally {
      await store.cleanup();
    }
  },
} as const;
