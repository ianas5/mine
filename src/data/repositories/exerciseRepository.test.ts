/** @jest-environment node */
import { setDbForTesting } from '@/core/db';

import { SEED_EXERCISES } from '../seed/exercises';
import { seedDatabase } from '../seed/seedDatabase';
import { createTestDb, type TestDb } from '../testing/createTestDb';
import { DuplicateExerciseNameError, exerciseRepository } from './exerciseRepository';

describe('exerciseRepository (real SQLite)', () => {
  let testDb: TestDb;

  beforeEach(() => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
  });

  afterEach(() => testDb.close());

  it('seeds the full library and is idempotent on re-run', async () => {
    await seedDatabase();
    await seedDatabase();

    const active = await exerciseRepository.listActive();
    expect(active).toHaveLength(SEED_EXERCISES.length);
    expect(active.length).toBeGreaterThanOrEqual(100);
  });

  it('sorts active exercises by name and excludes archived', async () => {
    await seedDatabase();
    const active = await exerciseRepository.listActive();

    const names = active.map((e) => e.name);
    expect(names).toEqual([...names].sort((a, b) => a.localeCompare(b)));
  });

  it('creates a custom exercise and returns it hydrated', async () => {
    const created = await exerciseRepository.createCustom({
      name: 'Cable Y-Raise',
      primaryMuscleGroup: 'shoulders',
      loadType: 'external',
      defaultUnilateral: false,
    });

    expect(created).toMatchObject({
      name: 'Cable Y-Raise',
      primaryMuscleGroup: 'shoulders',
      isCustom: true,
      isArchived: false,
    });
    expect(created.id).toMatch(/^ex_custom_/);
  });

  it('rejects a case-insensitively duplicate name', async () => {
    await exerciseRepository.createCustom({
      name: 'My Lift',
      primaryMuscleGroup: 'chest',
      loadType: 'external',
      defaultUnilateral: false,
    });

    await expect(
      exerciseRepository.createCustom({
        name: 'my lift',
        primaryMuscleGroup: 'back',
        loadType: 'external',
        defaultUnilateral: false,
      }),
    ).rejects.toBeInstanceOf(DuplicateExerciseNameError);
  });

  it('archives an exercise out of the active list and back', async () => {
    await seedDatabase();
    const before = await exerciseRepository.listActive();
    const target = before[0]!;

    await exerciseRepository.archive(target.id);
    const active = await exerciseRepository.listActive();
    const archived = await exerciseRepository.listArchived();
    expect(active.find((e) => e.id === target.id)).toBeUndefined();
    expect(archived.find((e) => e.id === target.id)).toBeDefined();

    await exerciseRepository.unarchive(target.id);
    expect((await exerciseRepository.listActive()).find((e) => e.id === target.id)).toBeDefined();
  });

  it('hard-deletes custom exercises but never seed rows (archive, not delete)', async () => {
    await seedDatabase();
    const seedId = SEED_EXERCISES.map((e) => `ex_seed_${e.slug}`)[0]!;
    const custom = await exerciseRepository.createCustom({
      name: 'Disposable',
      primaryMuscleGroup: 'core',
      loadType: 'timed',
      defaultUnilateral: false,
    });

    await exerciseRepository.remove(seedId); // no-op: not custom
    await exerciseRepository.remove(custom.id); // deleted: custom

    const all = [
      ...(await exerciseRepository.listActive()),
      ...(await exerciseRepository.listArchived()),
    ];
    expect(all.find((e) => e.id === seedId)).toBeDefined();
    expect(all.find((e) => e.id === custom.id)).toBeUndefined();
  });
});
