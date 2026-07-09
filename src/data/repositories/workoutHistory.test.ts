/** @jest-environment node */
import { setDbForTesting } from '@/core/db';

import { seedDatabase } from '../seed/seedDatabase';
import { createTestDb, type TestDb } from '../testing/createTestDb';
import { workoutRepository, type NewWorkoutInput } from './workoutRepository';

const BENCH = 'ex_seed_barbell-bench-press';

const session = (weightKg: number, reps: number): NewWorkoutInput => ({
  name: 'Push',
  startedAt: null,
  endedAt: null,
  notes: null,
  exercises: [
    {
      exerciseId: BENCH,
      unilateralCounting: 'none',
      notes: null,
      sets: [
        { weightKg, reps, rpe: null, warmup: false },
        { weightKg, reps: reps - 1, rpe: null, warmup: false },
      ],
    },
  ],
});

describe('workout history & prefill (real SQLite)', () => {
  let testDb: TestDb;

  beforeEach(async () => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
    await seedDatabase();
  });

  afterEach(() => testDb.close());

  it('has no preview for a never-performed exercise (P8)', async () => {
    const preview = await workoutRepository.getExercisePreview(BENCH, 'external');
    expect(preview.last).toBeNull();
    expect(preview.bestE1rmKg).toBeNull();
  });

  it('previews Last from the most recent workout across sessions', async () => {
    await workoutRepository.saveCompletedWorkout(session(70, 10));
    await workoutRepository.saveCompletedWorkout(session(80, 8));

    const preview = await workoutRepository.getExercisePreview(BENCH, 'external');
    expect(preview.last?.sets).toEqual([
      { weightKg: 80, reps: 8 },
      { weightKg: 80, reps: 7 },
    ]);
    expect(preview.bestWeightSet?.weightKg).toBe(80);
  });

  it('recomputes the preview after editing a saved set (edit A → B sees it)', async () => {
    const id = await workoutRepository.saveCompletedWorkout(session(80, 8));
    const workout = await workoutRepository.getById(id);
    const firstSetId = workout!.exercises[0]!.sets[0]!.id;

    await workoutRepository.updateSet(firstSetId, { weightKg: 95 });

    const preview = await workoutRepository.getExercisePreview(BENCH, 'external');
    expect(preview.bestWeightSet?.weightKg).toBe(95);
  });

  it('recomputes (empties) the preview after deleting the workout', async () => {
    const id = await workoutRepository.saveCompletedWorkout(session(80, 8));
    await workoutRepository.remove(id);

    const preview = await workoutRepository.getExercisePreview(BENCH, 'external');
    expect(preview.last).toBeNull();
  });

  it('lists recent workouts newest first', async () => {
    await workoutRepository.saveCompletedWorkout({ ...session(70, 10), name: 'Older' });
    await workoutRepository.saveCompletedWorkout({ ...session(80, 8), name: 'Newer' });

    const recent = await workoutRepository.listRecent();
    expect(recent[0]?.name).toBe('Newer');
    expect(recent).toHaveLength(2);
  });
});
