/** @jest-environment node */
import { setDbForTesting } from '@/core/db';

import { createTestDb, type TestDb } from '../testing/createTestDb';
import { nutritionRepository, type TargetInput } from './nutritionRepository';

const T = (kcal: number, water: number | null = 3000): TargetInput => ({
  kcal,
  proteinG: 160,
  carbG: 200,
  fatG: 60,
  waterMl: water,
});

describe('nutrition target resolution & water (real SQLite)', () => {
  let testDb: TestDb;

  beforeEach(() => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
  });

  afterEach(() => testDb.close());

  it('returns null before the first target (insufficient-data, never a default)', async () => {
    await nutritionRepository.setTarget('2026-06-01', T(2200));
    expect(await nutritionRepository.resolveTargetForDate('2026-05-31')).toBeNull();
  });

  it('resolves the greatest effective_from ≤ date', async () => {
    await nutritionRepository.setTarget('2026-01-01', T(2000));
    await nutritionRepository.setTarget('2026-06-01', T(2400));

    expect((await nutritionRepository.resolveTargetForDate('2026-03-15'))?.kcal).toBe(2000);
    expect((await nutritionRepository.resolveTargetForDate('2026-06-01'))?.kcal).toBe(2400);
    expect((await nutritionRepository.resolveTargetForDate('2026-09-01'))?.kcal).toBe(2400);
  });

  it("P5: setting new targets forward never changes an old day's target", async () => {
    await nutritionRepository.setTarget('2026-01-01', T(2000));
    // An old day resolves the original era.
    expect((await nutritionRepository.resolveTargetForDate('2026-03-01'))?.kcal).toBe(2000);

    // New targets from June forward.
    await nutritionRepository.setTarget('2026-06-01', T(2400));

    // The old day is unchanged; only dates from June forward see the new target.
    expect((await nutritionRepository.resolveTargetForDate('2026-03-01'))?.kcal).toBe(2000);
    expect((await nutritionRepository.resolveTargetForDate('2026-06-15'))?.kcal).toBe(2400);
  });

  it('upserts an era in place when the same effective_from is set again', async () => {
    await nutritionRepository.setTarget('2026-01-01', T(2000));
    await nutritionRepository.setTarget('2026-01-01', T(2100));

    expect(await nutritionRepository.listTargets()).toHaveLength(1);
    expect((await nutritionRepository.resolveTargetForDate('2026-02-01'))?.kcal).toBe(2100);
  });

  it('distinguishes a logged 0 from an unlogged day for water (§2.4)', async () => {
    expect(await nutritionRepository.getWater('2026-07-09')).toBeNull(); // unlogged

    await nutritionRepository.setWater('2026-07-09', 0);
    expect(await nutritionRepository.getWater('2026-07-09')).toBe(0); // a real logged zero

    await nutritionRepository.addWater('2026-07-09', 500);
    expect(await nutritionRepository.getWater('2026-07-09')).toBe(500);

    await nutritionRepository.addWater('2026-07-09', -900); // floored at 0
    expect(await nutritionRepository.getWater('2026-07-09')).toBe(0);
  });
});
