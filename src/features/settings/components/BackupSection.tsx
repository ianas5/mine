import { useState, type ReactNode } from 'react';
import { Text } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, Dialog, Section, showToast } from '@/core/ui';
import { backupService, ImportError, type ImportErrorCode } from '@/data/backup';

type Busy = 'idle' | 'export' | 'import';

function importErrorMessage(code: ImportErrorCode): string {
  switch (code) {
    case 'unreadable-archive':
      return "Couldn't open that file — pick a fitness backup .zip.";
    case 'invalid-data':
      return 'That backup is corrupted or not a fitness backup.';
    case 'unsupported-format':
      return 'That backup was made by an incompatible app version.';
    case 'schema-too-new':
      return 'That backup is from a newer app version — update the app first.';
    case 'unsupported-schema':
      return 'That backup is too old to import.';
    case 'aborted-no-safety':
      return 'Import cancelled — your data is unchanged.';
  }
}

/**
 * Data ownership (DATABASE §6): export a full zip via the share sheet, or import one
 * to replace all data. Import is guarded twice — a destructive "replace everything"
 * confirmation, and, if the automatic safety backup fails, an explicit
 * continue-without-safety confirmation (§6.7). Existing data is never touched until
 * validation passes and, on safety failure, the user confirms.
 */
export function BackupSection(): ReactNode {
  const theme = useTheme();
  const [busy, setBusy] = useState<Busy>('idle');
  const [pendingUri, setPendingUri] = useState<string | null>(null);
  const [safetyResolver, setSafetyResolver] = useState<((proceed: boolean) => void) | null>(null);

  const onExport = async (): Promise<void> => {
    setBusy('export');
    try {
      await backupService.exportAndShare();
    } catch {
      showToast('Could not create the backup');
    } finally {
      setBusy('idle');
    }
  };

  const onImportPress = async (): Promise<void> => {
    try {
      const uri = await backupService.pickArchive();
      if (uri) setPendingUri(uri);
    } catch {
      showToast('Could not open the file picker');
    }
  };

  const runImport = async (uri: string): Promise<void> => {
    setPendingUri(null);
    setBusy('import');
    try {
      await backupService.importArchive(uri, {
        // Resolve this promise from the safety-failed Dialog buttons below.
        onSafetyExportFailed: () =>
          new Promise<boolean>((resolve) => setSafetyResolver(() => resolve)),
      });
      showToast('Data restored', 'success');
    } catch (err) {
      if (err instanceof ImportError) {
        showToast(importErrorMessage(err.code));
      } else {
        showToast('Import failed — your data is unchanged');
      }
    } finally {
      setSafetyResolver(null);
      setBusy('idle');
    }
  };

  return (
    <Section title="Backup">
      <Card style={{ gap: theme.space.md }}>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
          Your data lives only on this device. Export a backup you own, or import one to restore —
          importing replaces everything currently in the app.
        </Text>
        <Button
          label="Export backup"
          loading={busy === 'export'}
          disabled={busy !== 'idle'}
          onPress={() => void onExport()}
        />
        <Button
          label="Import backup"
          variant="secondary"
          loading={busy === 'import'}
          disabled={busy !== 'idle'}
          onPress={() => void onImportPress()}
        />
      </Card>

      <Dialog
        visible={pendingUri !== null}
        title="Replace all data?"
        message="This overwrites everything currently in the app with the backup's contents. A safety backup of your current data is created first."
        confirmLabel="Replace"
        onConfirm={() => {
          if (pendingUri) void runImport(pendingUri);
        }}
        onCancel={() => setPendingUri(null)}
      />

      <Dialog
        visible={safetyResolver !== null}
        title="Safety backup failed"
        message="We couldn't create a safety backup of your current data. Continue with the import anyway? Your current data will be replaced."
        confirmLabel="Continue"
        cancelLabel="Cancel"
        onConfirm={() => safetyResolver?.(true)}
        onCancel={() => safetyResolver?.(false)}
      />
    </Section>
  );
}
