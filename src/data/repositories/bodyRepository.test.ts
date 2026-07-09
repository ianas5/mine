/** @jest-environment node */
import { setDbForTesting } from '@/core/db';

import { createTestDb, type TestDb } from '../testing/createTestDb';
import { bodyRepository } from './bodyRepository';

const DATE = '2026-07-09';

describe('bodyRepository merge-upsert (real SQLite)', () => {
  let testDb: TestDb;

  beforeEach(() => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
  });

  afterEach(() => testDb.close());

  it('creates a partial snapshot (absent fields default to null)', async () => {
    await bodyRepository.saveSnapshot(DATE, { weightKg: 80 });
    const snap = await bodyRepository.getSnapshot(DATE);
    expect(snap?.weightKg).toBe(80);
    expect(snap?.waistCm).toBeNull();
  });

  it('merges a same-date save: omitted fields keep their stored values (omit ≠ clear)', async () => {
    await bodyRepository.saveSnapshot(DATE, { weightKg: 80, waistCm: 85 });
    // A later save that omits waist must NOT erase it.
    await bodyRepository.saveSnapshot(DATE, { weightKg: 79.5 });

    const snap = await bodyRepository.getSnapshot(DATE);
    expect(snap?.weightKg).toBe(79.5); // updated
    expect(snap?.waistCm).toBe(85); // preserved, not nulled
  });

  it('stores bilateral sites per side, never collapsed', async () => {
    await bodyRepository.saveSnapshot(DATE, { leftArmCm: 38, rightArmCm: 38.5 });
    const snap = await bodyRepository.getSnapshot(DATE);
    expect(snap?.leftArmCm).toBe(38);
    expect(snap?.rightArmCm).toBe(38.5);
  });

  it('clears exactly one field only via the explicit clear action', async () => {
    await bodyRepository.saveSnapshot(DATE, { weightKg: 80, waistCm: 85, chestCm: 100 });

    await bodyRepository.clearField(DATE, 'waistCm');

    const snap = await bodyRepository.getSnapshot(DATE);
    expect(snap?.waistCm).toBeNull(); // cleared
    expect(snap?.weightKg).toBe(80); // untouched
    expect(snap?.chestCm).toBe(100); // untouched
  });

  it('builds the weight log newest-first from weigh-ins only', async () => {
    await bodyRepository.saveSnapshot('2026-07-01', { weightKg: 82 });
    await bodyRepository.saveSnapshot('2026-07-09', { weightKg: 80 });
    await bodyRepository.saveSnapshot('2026-07-05', { waistCm: 85 }); // no weight → excluded

    const log = await bodyRepository.getWeightLog();
    expect(log.map((p) => p.date)).toEqual(['2026-07-09', '2026-07-01']);
    expect(log[0]?.weightKg).toBe(80);
  });
});
