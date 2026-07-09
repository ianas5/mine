import type { Exercise } from '@/domain/models';

import { useSessionStore } from '../stores/useSessionStore';
import { parseSessionDraft, serializeSession } from './sessionDraftSchema';

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

describe('sessionDraftSchema', () => {
  beforeEach(() => state().actions.discard());

  it('round-trips a live session through serialize → parse', () => {
    state().actions.start(1710000000000, 'Push Day');
    state().actions.addExercise(exercise(), [
      { weightKg: 80, reps: 8 },
      { weightKg: 80, reps: 8 },
    ]);
    const exId = state().exercises[0]!.localId;
    const setId = state().exercises[0]!.sets[0]!.localId;
    state().actions.toggleSetDone(exId, setId);

    const draft = parseSessionDraft(serializeSession(state()));

    expect(draft).not.toBeNull();
    expect(draft?.name).toBe('Push Day');
    expect(draft?.startedAt).toBe(1710000000000);
    expect(draft?.exercises[0]?.exerciseId).toBe('ex_seed_bench');
    expect(draft?.exercises[0]?.sets[0]).toEqual({
      weightKg: 80,
      reps: 8,
      rpe: null,
      warmup: false,
      done: true,
    });
  });

  it('returns null for unparseable JSON (discarded gracefully)', () => {
    expect(parseSessionDraft('{ not json')).toBeNull();
  });

  it('returns null for a wrong-shape or legacy-version payload', () => {
    expect(parseSessionDraft(JSON.stringify({ version: 1, name: 'x' }))).toBeNull();
    expect(
      parseSessionDraft(JSON.stringify({ version: 2, name: 'x', startedAt: 0, exercises: [] })),
    ).toBeNull();
  });
});
