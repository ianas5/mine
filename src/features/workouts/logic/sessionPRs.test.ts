/** @jest-environment node */
import { setDbForTesting } from '@/core/db';
import { workoutRepository, type NewSetInput } from '@/data/repositories/workoutRepository';
import { seedDatabase } from '@/data/seed/seedDatabase';
import { createTestDb, type TestDb } from '@/data/testing/createTestDb';
import { computeExerciseReport } from '@/domain/analytics';
import { computeExerciseBests } from '@/domain/fitness';
import type { Exercise } from '@/domain/models';

import { useSessionStore } from '../stores/useSessionStore';
import { computeSessionPRs } from './sessionPRs';

const BENCH = 'ex_seed_barbell-bench-press';

const bench = (): Exercise => ({
  id: BENCH,
  name: 'Barbell Bench Press',
  primaryMuscleGroup: 'chest',
  secondaryMuscleGroups: [],
  loadType: 'external',
  defaultUnilateral: false,
  isCustom: false,
  isArchived: false,
  notes: null,
});

const store = () => useSessionStore.getState();

const saveWorkout = (sets: readonly { weightKg: number; reps: number }[]): Promise<string> =>
  workoutRepository.saveCompletedWorkout({
    name: 'Push',
    startedAt: null,
    endedAt: null,
    notes: null,
    exercises: [
      {
        exerciseId: BENCH,
        unilateralCounting: 'none',
        notes: null,
        sets: sets.map((s): NewSetInput => ({ ...s, rpe: null, warmup: false })),
      },
    ],
  });

const sessionOf = (weightKg: number, reps: number): void => {
  store().actions.discard();
  store().actions.start(1710000000000, 'Push');
  store().actions.addExercise(bench(), [{ weightKg, reps }]);
};

describe('session PRs & report (real SQLite)', () => {
  let testDb: TestDb;

  beforeEach(async () => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
    await seedDatabase();
    store().actions.discard();
  });

  afterEach(() => {
    store().actions.discard();
    testDb.close();
  });

  it('detects a new PR the session beats vs saved history', async () => {
    await saveWorkout([{ weightKg: 100, reps: 5 }]);
    sessionOf(105, 5);

    const prs = await computeSessionPRs(store(), null);
    expect(prs.totalCount).toBeGreaterThan(0);
    expect(prs.groups[0]?.events.map((e) => e.kind)).toEqual(
      expect.arrayContaining(['weight', 'e1rm']),
    );
  });

  it('reports no PRs when the session only ties history (strictly greater rule)', async () => {
    await saveWorkout([{ weightKg: 100, reps: 5 }]);
    sessionOf(100, 5);

    expect((await computeSessionPRs(store(), null)).totalCount).toBe(0);
  });

  it('recedes the record when the record-holding workout is deleted', async () => {
    await saveWorkout([{ weightKg: 100, reps: 5 }]);
    const record = await saveWorkout([{ weightKg: 120, reps: 3 }]);

    const before = await workoutRepository.getExerciseSetHistory(BENCH);
    expect(computeExerciseBests(before!.rows, 'external', null).heaviestWeightKg).toBe(120);

    await workoutRepository.remove(record);

    const after = await workoutRepository.getExerciseSetHistory(BENCH);
    expect(computeExerciseBests(after!.rows, 'external', null).heaviestWeightKg).toBe(100);
  });

  it('produces a report whose numbers reconcile with the saved history', async () => {
    await saveWorkout([
      { weightKg: 100, reps: 10 },
      { weightKg: 100, reps: 8 },
    ]);
    const history = await workoutRepository.getExerciseSetHistory(BENCH);
    const report = computeExerciseReport(history!.rows, 'external', null);

    expect(report.totalSessions).toBe(1);
    expect(report.totalWorkingSets).toBe(2);
    expect(report.totalVolumeKg).toBe(100 * 10 + 100 * 8);
    expect(report.bests.heaviestWeightKg).toBe(100);
  });
});
