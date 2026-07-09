import type { MuscleGroup, LoadType, UnilateralCounting } from '@/domain/fitness';

import type { TrainingExercise, TrainingSet, TrainingWorkout } from './trainingData';

/**
 * A hand-computed 4-week training fixture (ANALYTICS §5.1/§5.6 acceptance). Four Monday
 * sessions ending in the current week; volumes are simple w×r×sets so every metric can
 * be reconciled by hand in the tests. Not a `.test` file — imported by the calculators'
 * suites. `today` = 2026-03-15 (Sunday); its ISO week starts 2026-03-09 (Monday).
 */

const set = (weightKg: number, reps: number, warmup = false): TrainingSet => ({
  weightKg,
  reps,
  warmup,
});

const ex = (
  exerciseId: string,
  name: string,
  primaryMuscleGroup: MuscleGroup,
  sets: readonly TrainingSet[],
  loadType: LoadType = 'external',
  counting: UnilateralCounting = 'none',
): TrainingExercise => ({ exerciseId, name, primaryMuscleGroup, loadType, counting, sets });

export const FIXTURE_TODAY = '2026-03-15';

export function sampleTrainingWorkouts(): TrainingWorkout[] {
  return [
    {
      id: 'w1',
      date: '2026-02-16',
      startedAt: 0,
      endedAt: 3_600_000, // 60 min — the only session with both timestamps
      exercises: [ex('bench', 'Bench Press', 'chest', [set(100, 5), set(100, 5), set(100, 5)])],
    },
    {
      id: 'w2',
      date: '2026-02-23',
      startedAt: null,
      endedAt: null,
      exercises: [
        ex('bench', 'Bench Press', 'chest', [set(102, 5), set(102, 5), set(102, 5)]),
        ex('row', 'Barbell Row', 'back', [set(60, 10), set(60, 10), set(60, 10)]),
      ],
    },
    {
      id: 'w3',
      date: '2026-03-02',
      startedAt: null,
      endedAt: null,
      exercises: [
        ex('bench', 'Bench Press', 'chest', [set(104, 5), set(104, 5), set(104, 5)]),
        ex('squat', 'Back Squat', 'quads', [set(140, 5), set(140, 5), set(140, 5)]),
      ],
    },
    {
      id: 'w4',
      date: '2026-03-09', // current week (Monday)
      startedAt: null,
      endedAt: null,
      exercises: [
        ex('bench', 'Bench Press', 'chest', [set(106, 5), set(106, 5), set(106, 5)]),
        ex('row', 'Barbell Row', 'back', [set(62, 10), set(62, 10), set(62, 10)]),
      ],
    },
  ];
}

// Volume by group (external, no doubling):
//   chest  = (100+102+104+106) × 5 × 3 = 6180
//   back   = (60+62) × 10 × 3          = 3660
//   quads  = 140 × 5 × 3               = 2100
//   total                             = 11940
export const FIXTURE = {
  totalVolumeKg: 11940,
  chestVolumeKg: 6180,
  backVolumeKg: 3660,
  quadsVolumeKg: 2100,
} as const;
