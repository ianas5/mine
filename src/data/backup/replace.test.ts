/** @jest-environment node */
import { setDbForTesting } from '@/core/db';

import { sets } from '../schema/tables';
import { createTestDb, type TestDb } from '../testing/createTestDb';
import type { BackupData } from './backupSchema';
import { collectBackupData } from './collect';
import { sampleBackupData } from './backupTestKit';
import { replaceAllData } from './replace';

let testDb: TestDb;

beforeEach(() => {
  testDb = createTestDb();
  setDbForTesting(testDb.db);
});

afterEach(() => {
  testDb.close();
});

describe('replaceAllData', () => {
  it('inserts a full dataset across all tables in FK order', async () => {
    await replaceAllData(sampleBackupData());
    const round = await collectBackupData();
    expect(round).toEqual(sampleBackupData());
  });

  it('replaces (not merges) prior data', async () => {
    await replaceAllData(sampleBackupData());

    const next: BackupData = {
      ...sampleBackupData(),
      exercises: [
        {
          id: 'ex2',
          name: 'Squat',
          primaryMuscleGroup: 'quads',
          secondaryMuscleGroups: '[]',
          loadType: 'external',
          defaultUnilateral: 0,
          isCustom: 0,
          isArchived: 0,
          notes: null,
          createdAt: 2,
          updatedAt: 2,
        },
      ],
      // Drop the child rows that referenced the now-absent ex1 so the FK holds.
      templateExercises: [],
      workoutExercises: [],
      sets: [],
    };
    await replaceAllData(next);

    const round = await collectBackupData();
    expect(round.exercises.map((e) => e.id)).toEqual(['ex2']);
    expect(round.sets).toHaveLength(0);
  });

  it('rolls back the whole replace (deletes included) when an insert fails mid-transaction', async () => {
    await replaceAllData(sampleBackupData());
    const before = await collectBackupData();

    // Inject a failure at the `sets` insert — after every table has been deleted and
    // several re-inserted. The rollback must restore ALL of it, deletes included.
    // (A JS-layer throw, deterministic across Jest's shared better-sqlite3 realm.)
    const realInsert = testDb.db.insert.bind(testDb.db);
    const spy = jest.spyOn(testDb.db, 'insert').mockImplementation((table) => {
      if (table === sets) throw new Error('injected mid-transaction failure');
      return realInsert(table);
    });

    await expect(replaceAllData(sampleBackupData())).rejects.toThrow('injected');
    spy.mockRestore();

    const after = await collectBackupData();
    expect(after).toEqual(before); // transaction rolled the whole replace back
  });
});
