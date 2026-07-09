/** @jest-environment node */
import { setDbForTesting } from '@/core/db';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import { seedDatabase } from '@/data/seed/seedDatabase';
import { createTestDb, type TestDb } from '@/data/testing/createTestDb';
import type { Exercise } from '@/domain/models';

import { serializeSession } from '../schemas/sessionDraftSchema';
import { useSessionStore, type SessionState } from '../stores/useSessionStore';
import { sessionToWorkoutInput } from './sessionMapping';
import { recoverSession } from './sessionRecovery';

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

/** Compare sessions by value (localIds are regenerated on restore, so exclude them). */
const shape = (state: SessionState) => ({
  name: state.name,
  startedAt: state.startedAt,
  exercises: state.exercises.map((ex) => ({
    exerciseId: ex.exerciseId,
    name: ex.name,
    loadType: ex.loadType,
    unilateralCounting: ex.unilateralCounting,
    sets: ex.sets.map((s) => ({
      weightKg: s.weightKg,
      reps: s.reps,
      rpe: s.rpe,
      warmup: s.warmup,
      done: s.done,
    })),
  })),
});

describe('session crash recovery (real SQLite)', () => {
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

  it('restores the exact in-progress session after a simulated force-kill', async () => {
    // Build and log a partial session.
    store().actions.start(1710000000000, 'Push Day');
    store().actions.addExercise(bench(), [
      { weightKg: 80, reps: 8 },
      { weightKg: 80, reps: 8 },
    ]);
    const exId = store().exercises[0]!.localId;
    store().actions.toggleSetDone(exId, store().exercises[0]!.sets[0]!.localId);
    const before = shape(store());

    // Checkpoint, then wipe memory — the crash.
    await workoutRepository.checkpointDraft(serializeSession(store()));
    store().actions.discard();
    expect(store().active).toBe(false);

    // Relaunch recovery.
    const outcome = await recoverSession();

    expect(outcome).toBe('restored');
    expect(store().active).toBe(true);
    expect(store().recovered).toBe(true);
    expect(shape(store())).toEqual(before);
  });

  it('deletes the draft transactionally when the workout is finished', async () => {
    store().actions.start(1710000000000, 'Push Day');
    store().actions.addExercise(bench());
    const exId = store().exercises[0]!.localId;
    store().actions.updateSet(exId, store().exercises[0]!.sets[0]!.localId, {
      weightKg: 80,
      reps: 8,
    });
    await workoutRepository.checkpointDraft(serializeSession(store()));
    expect(await workoutRepository.loadDraft()).not.toBeNull();

    await workoutRepository.saveCompletedWorkout(sessionToWorkoutInput(store(), 1710000600000));

    expect(await workoutRepository.loadDraft()).toBeNull();
  });

  it('discards a corrupt draft gracefully (no crash, no restore)', async () => {
    await workoutRepository.checkpointDraft('{ corrupt payload');

    const outcome = await recoverSession();

    expect(outcome).toBe('discarded');
    expect(store().active).toBe(false);
    expect(await workoutRepository.loadDraft()).toBeNull();
  });

  it('discards an empty session draft rather than recovering nothing', async () => {
    store().actions.start(1710000000000, 'Empty');
    await workoutRepository.checkpointDraft(serializeSession(store()));
    store().actions.discard();

    expect(await recoverSession()).toBe('discarded');
    expect(store().active).toBe(false);
  });

  it('checkpoint upserts a single row (id = 1) rather than accumulating drafts', async () => {
    await workoutRepository.checkpointDraft('first');
    await workoutRepository.checkpointDraft('second');

    expect(await workoutRepository.loadDraft()).toBe('second');

    await workoutRepository.discardDraft();
    expect(await workoutRepository.loadDraft()).toBeNull();
  });
});
