/** @jest-environment node */
import { setDbForTesting } from '@/core/db';
import { todayIso } from '@/core/utils';

import { exerciseRepository } from './exerciseRepository';
import { seedDatabase } from '../seed/seedDatabase';
import { createTestDb, type TestDb } from '../testing/createTestDb';
import { workoutRepository, type NewWorkoutInput } from './workoutRepository';

const SEED_ID = 'ex_seed_barbell-bench-press';

const workoutWith = (exerciseId: string): NewWorkoutInput => ({
  name: 'Push Day',
  startedAt: 1_000,
  endedAt: 4_600_000,
  notes: null,
  exercises: [
    {
      exerciseId,
      unilateralCounting: 'none',
      notes: null,
      sets: [
        { weightKg: 60, reps: 10, rpe: null, warmup: true },
        { weightKg: 80, reps: 8, rpe: 8, warmup: false },
        { weightKg: 0, reps: 0, rpe: null, warmup: false }, // empty → dropped
      ],
    },
  ],
});

describe('workoutRepository (real SQLite)', () => {
  let testDb: TestDb;

  beforeEach(async () => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
    await seedDatabase();
  });

  afterEach(() => testDb.close());

  it('persists a workout tree and drops empty sets', async () => {
    const id = await workoutRepository.saveCompletedWorkout(workoutWith(SEED_ID));
    const workout = await workoutRepository.getById(id);

    expect(workout?.exercises).toHaveLength(1);
    expect(workout?.exercises[0]?.sets).toHaveLength(2); // warm-up kept, empty dropped
    expect(workout?.exercises[0]?.sets[0]?.warmup).toBe(true);
    expect(workout?.exercises[0]?.sets[1]).toMatchObject({ weightKg: 80, reps: 8, rpe: 8 });
    expect(workout?.exercises[0]?.name).toBe('Barbell Bench Press');
  });

  it('allows two workouts on the same date (edge 3)', async () => {
    await workoutRepository.saveCompletedWorkout(workoutWith(SEED_ID));
    await workoutRepository.saveCompletedWorkout(workoutWith(SEED_ID));
    expect(await workoutRepository.countOnDate(todayIso())).toBe(2);
  });

  it('rolls back entirely when a referenced exercise does not exist', async () => {
    await expect(
      workoutRepository.saveCompletedWorkout(workoutWith('ex_does_not_exist')),
    ).rejects.toBeDefined();
    expect(await workoutRepository.countOnDate(todayIso())).toBe(0);
  });

  it('FK RESTRICT prevents deleting an exercise referenced by history', async () => {
    const custom = await exerciseRepository.createCustom({
      name: 'Referenced Lift',
      primaryMuscleGroup: 'chest',
      loadType: 'external',
      defaultUnilateral: false,
    });
    await workoutRepository.saveCompletedWorkout(workoutWith(custom.id));

    await expect(exerciseRepository.remove(custom.id)).rejects.toBeDefined();
    const stillThere = (await exerciseRepository.listActive()).find((e) => e.id === custom.id);
    expect(stillThere).toBeDefined();
  });
});
