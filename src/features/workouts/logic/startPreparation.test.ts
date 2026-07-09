import type { Workout } from '@/domain/models';

import { prepareRepeatLast } from './startPreparation';

const workout = (): Workout => ({
  id: 'wk_1',
  date: '2026-07-08',
  name: 'Push A',
  startedAt: null,
  endedAt: null,
  notes: null,
  exercises: [
    {
      id: 'we_1',
      exerciseId: 'ex_seed_barbell-bench-press',
      name: 'Barbell Bench Press',
      loadType: 'external',
      unilateralCounting: 'none',
      position: 0,
      notes: null,
      sets: [
        {
          id: 's1',
          position: 0,
          weightKg: 60,
          reps: 10,
          rpe: null,
          rir: null,
          warmup: true,
          notes: null,
        },
        {
          id: 's2',
          position: 1,
          weightKg: 100,
          reps: 5,
          rpe: null,
          rir: null,
          warmup: false,
          notes: null,
        },
        {
          id: 's3',
          position: 2,
          weightKg: 100,
          reps: 5,
          rpe: null,
          rir: null,
          warmup: false,
          notes: null,
        },
      ],
    },
  ],
});

describe('prepareRepeatLast', () => {
  it('reloads only the working sets as pre-fill and keeps the workout untouched', () => {
    const { name, prepared } = prepareRepeatLast(workout());
    expect(name).toBe('Push A');
    expect(prepared).toHaveLength(1);
    expect(prepared[0]?.exerciseId).toBe('ex_seed_barbell-bench-press');
    expect(prepared[0]?.target).toBeNull();
    // The warm-up (60×10) is dropped; the two working sets carry over.
    expect(prepared[0]?.sets).toEqual([
      { weightKg: 100, reps: 5 },
      { weightKg: 100, reps: 5 },
    ]);
  });
});
