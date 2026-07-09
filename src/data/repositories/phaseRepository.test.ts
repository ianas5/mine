/** @jest-environment node */
import { setDbForTesting } from '@/core/db';

import { createTestDb, type TestDb } from '../testing/createTestDb';
import { PhaseValidationError, phaseRepository, type PhaseInput } from './phaseRepository';

let testDb: TestDb;

beforeEach(() => {
  testDb = createTestDb();
  setDbForTesting(testDb.db);
});

afterEach(() => {
  testDb.close();
});

const base = (over: Partial<PhaseInput> = {}): PhaseInput => ({
  name: 'Cut',
  type: 'cutting',
  startDate: '2026-01-01',
  endDate: '2026-02-28',
  notes: null,
  ...over,
});

describe('phaseRepository — CRUD', () => {
  it('creates and reads back a phase', async () => {
    const id = await phaseRepository.createPhase(base());
    const phase = await phaseRepository.getPhase(id);
    expect(phase).toMatchObject({ name: 'Cut', type: 'cutting', endDate: '2026-02-28' });
  });

  it('trims a blank name to a fallback', async () => {
    const id = await phaseRepository.createPhase(base({ name: '   ' }));
    expect((await phaseRepository.getPhase(id))!.name).toBe('Phase');
  });

  it('lists phases most-recent first', async () => {
    await phaseRepository.createPhase(base({ startDate: '2026-01-01', endDate: '2026-02-28' }));
    await phaseRepository.createPhase(base({ startDate: '2026-03-01', endDate: '2026-04-30' }));
    const list = await phaseRepository.listPhases();
    expect(list.map((p) => p.startDate)).toEqual(['2026-03-01', '2026-01-01']);
  });

  it('deletes a phase', async () => {
    const id = await phaseRepository.createPhase(base());
    await phaseRepository.deletePhase(id);
    expect(await phaseRepository.getPhase(id)).toBeNull();
  });
});

describe('phaseRepository — no-overlap invariant', () => {
  it('rejects a phase overlapping an existing closed phase', async () => {
    await phaseRepository.createPhase(base({ startDate: '2026-01-01', endDate: '2026-02-28' }));
    await expect(
      phaseRepository.createPhase(base({ startDate: '2026-02-15', endDate: '2026-03-31' })),
    ).rejects.toBeInstanceOf(PhaseValidationError);
  });

  it('allows adjacent phases (end day before next start)', async () => {
    await phaseRepository.createPhase(base({ startDate: '2026-01-01', endDate: '2026-02-28' }));
    await expect(
      phaseRepository.createPhase(base({ startDate: '2026-03-01', endDate: '2026-04-30' })),
    ).resolves.toBeTruthy();
  });

  it('rejects a second ongoing phase (both open-ended overlap)', async () => {
    await phaseRepository.createPhase(base({ startDate: '2026-01-01', endDate: null }));
    await expect(
      phaseRepository.createPhase(base({ startDate: '2026-06-01', endDate: null })),
    ).rejects.toMatchObject({ reason: 'overlap' });
  });

  it('rejects a new phase that runs into the ongoing one', async () => {
    await phaseRepository.createPhase(base({ startDate: '2026-03-01', endDate: null }));
    await expect(
      phaseRepository.createPhase(base({ startDate: '2026-01-01', endDate: '2026-03-15' })),
    ).rejects.toMatchObject({ reason: 'overlap' });
  });

  it('rejects an end date before the start date', async () => {
    await expect(
      phaseRepository.createPhase(base({ startDate: '2026-02-01', endDate: '2026-01-01' })),
    ).rejects.toMatchObject({ reason: 'end-before-start' });
  });
});

describe('phaseRepository — ongoing + lookup', () => {
  it('returns the single ongoing phase', async () => {
    await phaseRepository.createPhase(base({ startDate: '2026-01-01', endDate: '2026-02-28' }));
    const ongoingId = await phaseRepository.createPhase(
      base({ name: 'Bulk', type: 'lean_bulk', startDate: '2026-03-01', endDate: null }),
    );
    expect((await phaseRepository.getOngoingPhase())!.id).toBe(ongoingId);
  });

  it('resolves the phase active on a given date, or null in a gap', async () => {
    await phaseRepository.createPhase(base({ startDate: '2026-01-01', endDate: '2026-02-28' }));
    await phaseRepository.createPhase(
      base({ name: 'Bulk', type: 'lean_bulk', startDate: '2026-04-01', endDate: null }),
    );
    expect((await phaseRepository.getPhaseForDate('2026-02-01'))!.type).toBe('cutting');
    expect((await phaseRepository.getPhaseForDate('2026-05-01'))!.type).toBe('lean_bulk');
    expect(await phaseRepository.getPhaseForDate('2026-03-15')).toBeNull();
  });
});

describe('phaseRepository — end + update', () => {
  it('ends an ongoing phase (end-yesterday UX) and frees the slot', async () => {
    const id = await phaseRepository.createPhase(base({ startDate: '2026-01-01', endDate: null }));
    await phaseRepository.endPhase(id, '2026-02-28');
    expect((await phaseRepository.getPhase(id))!.endDate).toBe('2026-02-28');
    expect(await phaseRepository.getOngoingPhase()).toBeNull();
    // Slot is free: a new ongoing phase now fits.
    await expect(
      phaseRepository.createPhase(base({ startDate: '2026-03-01', endDate: null })),
    ).resolves.toBeTruthy();
  });

  it('re-checks overlap when a phase is edited to collide', async () => {
    await phaseRepository.createPhase(base({ startDate: '2026-01-01', endDate: '2026-02-28' }));
    const second = await phaseRepository.createPhase(
      base({ startDate: '2026-03-01', endDate: '2026-04-30' }),
    );
    await expect(
      phaseRepository.updatePhase(second, { startDate: '2026-02-01' }),
    ).rejects.toMatchObject({ reason: 'overlap' });
  });
});
