import type { Exercise } from '@/domain/models';

import { useSessionStore } from './useSessionStore';

const exercise = (overrides: Partial<Exercise> = {}): Exercise => ({
  id: 'ex_seed_bench',
  name: 'Bench Press',
  primaryMuscleGroup: 'chest',
  secondaryMuscleGroups: [],
  loadType: 'external',
  defaultUnilateral: false,
  isCustom: false,
  isArchived: false,
  notes: null,
  ...overrides,
});

const state = () => useSessionStore.getState();

describe('useSessionStore', () => {
  beforeEach(() => state().actions.discard());

  it('starts an empty active session', () => {
    state().actions.start(1000, 'Push Day');
    expect(state().active).toBe(true);
    expect(state().name).toBe('Push Day');
    expect(state().exercises).toHaveLength(0);
  });

  it('adds an exercise with one empty set and counting from defaultUnilateral', () => {
    state().actions.start(1000);
    state().actions.addExercise(exercise());
    state().actions.addExercise(
      exercise({ id: 'ex_uni', name: 'Cable Kickback', defaultUnilateral: true }),
    );

    expect(state().exercises[0]?.unilateralCounting).toBe('none');
    expect(state().exercises[0]?.sets).toHaveLength(1);
    expect(state().exercises[1]?.unilateralCounting).toBe('single_doubled');
  });

  it('prefills sets from history when adding an exercise', () => {
    state().actions.start(1000);
    state().actions.addExercise(exercise(), [
      { weightKg: 80, reps: 8 },
      { weightKg: 80, reps: 8 },
      { weightKg: 75, reps: 10 },
    ]);

    const sets = state().exercises[0]!.sets;
    expect(sets).toHaveLength(3);
    expect(sets[0]).toMatchObject({ weightKg: 80, reps: 8, done: false });
    expect(sets[2]).toMatchObject({ weightKg: 75, reps: 10 });
  });

  it('inherits the previous set values when adding a set', () => {
    state().actions.start(1000);
    state().actions.addExercise(exercise());
    const exId = state().exercises[0]!.localId;
    const setId = state().exercises[0]!.sets[0]!.localId;

    state().actions.updateSet(exId, setId, { weightKg: 80, reps: 8 });
    state().actions.addSet(exId);

    const sets = state().exercises[0]!.sets;
    expect(sets).toHaveLength(2);
    expect(sets[1]).toMatchObject({ weightKg: 80, reps: 8, done: false });
  });

  it('clamps set values to plausibility ranges', () => {
    state().actions.start(1000);
    state().actions.addExercise(exercise());
    const exId = state().exercises[0]!.localId;
    const setId = state().exercises[0]!.sets[0]!.localId;

    state().actions.updateSet(exId, setId, { weightKg: 99999, reps: -5 });
    expect(state().exercises[0]!.sets[0]).toMatchObject({ weightKg: 1000, reps: 0 });
  });

  it('toggles set done and discards the session', () => {
    state().actions.start(1000);
    state().actions.addExercise(exercise());
    const exId = state().exercises[0]!.localId;
    const setId = state().exercises[0]!.sets[0]!.localId;

    state().actions.toggleSetDone(exId, setId);
    expect(state().exercises[0]!.sets[0]!.done).toBe(true);

    state().actions.discard();
    expect(state().active).toBe(false);
    expect(state().exercises).toHaveLength(0);
  });
});
