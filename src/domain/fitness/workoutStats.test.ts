import { computeWorkoutStats, isCountableWorkout, type StatExercise } from './workoutStats';

const external = (sets: { weightKg: number; reps: number; warmup?: boolean }[]): StatExercise => ({
  loadType: 'external',
  unilateralCounting: 'none',
  sets: sets.map((s) => ({ weightKg: s.weightKg, reps: s.reps, warmup: s.warmup ?? false })),
});

describe('computeWorkoutStats (FITNESS_DOMAIN §3.5)', () => {
  it('returns zeros and is not countable for an empty workout (edge 2)', () => {
    expect(computeWorkoutStats([], null)).toEqual({
      workingSetCount: 0,
      totalVolumeKg: 0,
      volumeLowConfidence: false,
    });
    expect(isCountableWorkout([], null)).toBe(false);
  });

  it('counts only working sets and excludes warm-ups (edge 1)', () => {
    const stats = computeWorkoutStats(
      [
        external([
          { weightKg: 60, reps: 10, warmup: true },
          { weightKg: 80, reps: 8 },
          { weightKg: 80, reps: 8 },
        ]),
      ],
      null,
    );
    expect(stats.workingSetCount).toBe(2);
    expect(stats.totalVolumeKg).toBe(1280);
    expect(isCountableWorkout([external([{ weightKg: 80, reps: 8 }])], null)).toBe(true);
  });

  it('doubles volume for single-logged unilateral exercises (edge 6)', () => {
    const stats = computeWorkoutStats(
      [
        {
          loadType: 'external',
          unilateralCounting: 'single_doubled',
          sets: [{ weightKg: 20, reps: 10, warmup: false }],
        },
      ],
      null,
    );
    expect(stats.totalVolumeKg).toBe(400);
  });

  it('flags low-confidence volume for bodyweight work without a known bodyweight (edge 5)', () => {
    const stats = computeWorkoutStats(
      [
        {
          loadType: 'bodyweight',
          unilateralCounting: 'none',
          sets: [{ weightKg: 0, reps: 10, warmup: false }],
        },
      ],
      null,
    );
    expect(stats.workingSetCount).toBe(1);
    expect(stats.totalVolumeKg).toBe(0);
    expect(stats.volumeLowConfidence).toBe(true);
  });
});
