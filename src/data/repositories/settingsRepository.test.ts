/** @jest-environment node */
import { setDbForTesting, subscribeToTables } from '@/core/db';

import { createTestDb, type TestDb } from '../testing/createTestDb';
import { settingsRepository } from './settingsRepository';

describe('settingsRepository (real SQLite via migrations)', () => {
  let testDb: TestDb;

  beforeEach(() => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
  });

  afterEach(() => {
    testDb.close();
  });

  it('migrates from empty and creates the default row on first get', async () => {
    const settings = await settingsRepository.get();

    expect(settings).toEqual({
      weeklyWorkoutTarget: 4,
      defaultBodyweightKg: null,
      heightCm: null,
      waterCupMl: 250,
    });
  });

  it('persists updates and merges partial patches', async () => {
    await settingsRepository.update({ weeklyWorkoutTarget: 5 });
    await settingsRepository.update({ heightCm: 178 });

    const settings = await settingsRepository.get();

    expect(settings.weeklyWorkoutTarget).toBe(5);
    expect(settings.heightCm).toBe(178);
    expect(settings.waterCupMl).toBe(250);
  });

  it('remains a single row: get never duplicates and id ≠ 1 is rejected by CHECK', async () => {
    await settingsRepository.get();
    await settingsRepository.get();

    const count = testDb.sqlite.prepare('SELECT COUNT(*) AS n FROM settings').get() as {
      n: number;
    };
    expect(count.n).toBe(1);

    expect(() =>
      testDb.sqlite
        .prepare(
          'INSERT INTO settings (id, weekly_workout_target, water_cup_ml, created_at, updated_at) VALUES (2, 4, 250, 0, 0)',
        )
        .run(),
    ).toThrow(/CHECK/i);
  });

  it('emits the settings change-bus event as the last step of update', async () => {
    const events: string[] = [];
    const unsubscribe = subscribeToTables(['settings'], () => events.push('settings'));

    await settingsRepository.update({ weeklyWorkoutTarget: 3 });
    unsubscribe();

    expect(events).toEqual(['settings']);
  });

  it('does not emit on reads', async () => {
    const events: string[] = [];
    const unsubscribe = subscribeToTables(['settings'], () => events.push('settings'));

    await settingsRepository.get();
    unsubscribe();

    expect(events).toEqual([]);
  });
});
