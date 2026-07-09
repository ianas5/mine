/** @jest-environment node */
import { sampleBackupData, makeEnvelopeJson } from './backupTestKit';
import { ImportError } from './importError';
import { parseAndUpgradeEnvelope } from './parseEnvelope';
import { CURRENT_SCHEMA_VERSION } from './schemaVersion';

const expectImportError = (fn: () => unknown, code: string): void => {
  try {
    fn();
    throw new Error('expected parseAndUpgradeEnvelope to throw');
  } catch (err) {
    expect(err).toBeInstanceOf(ImportError);
    expect((err as ImportError).code).toBe(code);
  }
};

describe('parseAndUpgradeEnvelope', () => {
  it('accepts a valid current-version envelope', () => {
    const envelope = parseAndUpgradeEnvelope(makeEnvelopeJson(sampleBackupData()));
    expect(envelope.schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
    expect(envelope.data.exercises).toHaveLength(1);
    expect(envelope.data.settings.weeklyWorkoutTarget).toBe(4);
  });

  it('strips unknown keys rather than failing', () => {
    const withExtra = JSON.parse(makeEnvelopeJson(sampleBackupData())) as Record<string, unknown>;
    (withExtra.data as { exercises: Record<string, unknown>[] }).exercises[0]!.bogus = 'x';
    const envelope = parseAndUpgradeEnvelope(JSON.stringify(withExtra));
    expect(envelope.data.exercises[0]).not.toHaveProperty('bogus');
  });

  it('rejects non-JSON text as invalid-data', () => {
    expectImportError(() => parseAndUpgradeEnvelope('not json at all'), 'invalid-data');
  });

  it('rejects a foreign app as invalid-data', () => {
    expectImportError(
      () =>
        parseAndUpgradeEnvelope(makeEnvelopeJson(sampleBackupData(), { app: 'some-other-app' })),
      'invalid-data',
    );
  });

  it('rejects a malformed data block as invalid-data', () => {
    const broken = JSON.parse(makeEnvelopeJson(sampleBackupData())) as Record<string, unknown>;
    (broken.data as { sets: Record<string, unknown>[] }).sets[0]!.reps = 'ten'; // wrong type
    expectImportError(() => parseAndUpgradeEnvelope(JSON.stringify(broken)), 'invalid-data');
  });

  it('refuses an unknown archive format as unsupported-format', () => {
    expectImportError(
      () => parseAndUpgradeEnvelope(makeEnvelopeJson(sampleBackupData(), { format: 999 })),
      'unsupported-format',
    );
  });

  it('refuses a newer schema version as schema-too-new', () => {
    expectImportError(
      () =>
        parseAndUpgradeEnvelope(
          makeEnvelopeJson(sampleBackupData(), { schemaVersion: CURRENT_SCHEMA_VERSION + 1 }),
        ),
      'schema-too-new',
    );
  });

  it('refuses an older schema with no upgrader path as unsupported-schema', () => {
    expectImportError(
      () =>
        parseAndUpgradeEnvelope(
          makeEnvelopeJson(sampleBackupData(), { schemaVersion: CURRENT_SCHEMA_VERSION - 1 }),
        ),
      'unsupported-schema',
    );
  });
});
