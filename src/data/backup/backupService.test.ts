/** @jest-environment node */
import { setDbForTesting } from '@/core/db';

import { setPhotoStore } from '../photos/photoStore';
import { sets } from '../schema/tables';
import { createTestDb, type TestDb } from '../testing/createTestDb';
import { setArchiveStore } from './archiveStore';
import { backupService } from './backupService';
import {
  createFakeArchiveStore,
  createFakePhotoStore,
  makeEnvelopeJson,
  sampleBackupData,
  type FakeArchiveControls,
} from './backupTestKit';
import { collectBackupData } from './collect';
import { replaceAllData } from './replace';
import { CURRENT_SCHEMA_VERSION } from './schemaVersion';

let testDb: TestDb;
let files: Map<string, string>;
let controls: FakeArchiveControls;

const alwaysConfirm = { onSafetyExportFailed: async () => true };
const neverConfirm = { onSafetyExportFailed: async () => false };

beforeEach(async () => {
  testDb = createTestDb();
  setDbForTesting(testDb.db);

  files = new Map<string, string>();
  setPhotoStore(createFakePhotoStore(files));
  const fake = createFakeArchiveStore(files);
  setArchiveStore(fake.store);
  controls = fake.controls;

  // A populated starting database with one real photo file on disk.
  await replaceAllData(sampleBackupData());
  files.set('2026-01-01_front_p1.jpg', 'original-bytes');
});

afterEach(() => {
  testDb.close();
});

describe('export', () => {
  it('packs a shareable archive of the current data + photos', async () => {
    await backupService.exportAndShare();
    expect(controls.lastExportUri).not.toBeNull();
    expect(controls.lastSharedUri).toBe(controls.lastExportUri);

    const archive = controls.archives.get(controls.lastExportUri!)!;
    expect(archive.photos.has('2026-01-01_front_p1.jpg')).toBe(true);
    const parsed = JSON.parse(archive.dataJson) as { schemaVersion: number; data: unknown };
    expect(parsed.schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
  });
});

describe('import — round trip', () => {
  it('export → mutate → import restores byte-equivalent data + photos', async () => {
    const before = await collectBackupData();
    await backupService.exportAndShare();
    const archiveUri = controls.lastExportUri!;

    // Heavily mutate the live DB + photos dir after the export.
    await replaceAllData({
      ...sampleBackupData(),
      settings: { ...sampleBackupData().settings, weeklyWorkoutTarget: 1 },
      exercises: [],
      templateExercises: [],
      workoutExercises: [],
      sets: [],
      progressPhotos: [],
    });
    files.set('stray_orphan.jpg', 'junk'); // a file with no row

    await backupService.importArchive(archiveUri, alwaysConfirm);

    const after = await collectBackupData();
    expect(after).toEqual(before);
    // Reconcile: imported photo present, the orphan file swept away.
    expect([...files.keys()]).toEqual(['2026-01-01_front_p1.jpg']);
  });
});

describe('import — refusals leave data untouched', () => {
  it('rejects a malformed archive as invalid-data and touches nothing', async () => {
    const before = await collectBackupData();
    controls.put('archive://bad', 'not valid json');

    await expect(backupService.importArchive('archive://bad', alwaysConfirm)).rejects.toMatchObject(
      {
        code: 'invalid-data',
      },
    );

    expect(await collectBackupData()).toEqual(before);
  });

  it('refuses a newer schema version and touches nothing', async () => {
    const before = await collectBackupData();
    controls.put(
      'archive://newer',
      makeEnvelopeJson(sampleBackupData(), { schemaVersion: CURRENT_SCHEMA_VERSION + 1 }),
    );

    await expect(
      backupService.importArchive('archive://newer', alwaysConfirm),
    ).rejects.toMatchObject({ code: 'schema-too-new' });

    expect(await collectBackupData()).toEqual(before);
  });

  it('rolls back and never commits photos when the replace fails mid-transaction', async () => {
    const before = await collectBackupData();
    controls.put('archive://ok', makeEnvelopeJson(sampleBackupData()));

    // Fail the replace at the `sets` insert (deterministic JS-layer throw).
    const realInsert = testDb.db.insert.bind(testDb.db);
    const spy = jest.spyOn(testDb.db, 'insert').mockImplementation((table) => {
      if (table === sets) throw new Error('injected mid-transaction failure');
      return realInsert(table);
    });

    await expect(backupService.importArchive('archive://ok', alwaysConfirm)).rejects.toThrow();
    spy.mockRestore();

    expect(await collectBackupData()).toEqual(before); // replace rolled back
    expect(files.has('2026-01-01_front_p1.jpg')).toBe(true); // photos never reconciled
  });
});

describe('import — safety export gate (DATABASE §6.7)', () => {
  it('aborts untouched when the safety export fails and the user declines', async () => {
    const before = await collectBackupData();
    controls.failSafetyExport = true;
    controls.put('archive://ok', makeEnvelopeJson(sampleBackupData()));

    await expect(backupService.importArchive('archive://ok', neverConfirm)).rejects.toMatchObject({
      code: 'aborted-no-safety',
    });

    expect(await collectBackupData()).toEqual(before);
  });

  it('proceeds when the safety export fails but the user explicitly confirms', async () => {
    controls.failSafetyExport = true;
    const imported = {
      ...sampleBackupData(),
      settings: { ...sampleBackupData().settings, weeklyWorkoutTarget: 7 },
    };
    controls.put('archive://ok', makeEnvelopeJson(imported), new Map());

    await backupService.importArchive('archive://ok', alwaysConfirm);

    const after = await collectBackupData();
    expect(after.settings.weeklyWorkoutTarget).toBe(7);
  });

  it('never asks for confirmation when the safety export succeeds', async () => {
    let asked = false;
    controls.put('archive://ok', makeEnvelopeJson(sampleBackupData()));

    await backupService.importArchive('archive://ok', {
      onSafetyExportFailed: async () => {
        asked = true;
        return true;
      },
    });

    expect(asked).toBe(false);
  });
});
