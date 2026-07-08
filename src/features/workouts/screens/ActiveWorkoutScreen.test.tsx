import { fireEvent, render, screen, waitFor } from '@testing-library/react-native';

import { setDbForTesting } from '@/core/db';
import { createTestDb, type TestDb } from '@/data/testing/createTestDb';
import type { Exercise } from '@/domain/models';

import { useSessionStore } from '../stores/useSessionStore';
import { ActiveWorkoutScreen } from './ActiveWorkoutScreen';

const bench: Exercise = {
  id: 'ex_seed_barbell-bench-press',
  name: 'Barbell Bench Press',
  primaryMuscleGroup: 'chest',
  secondaryMuscleGroups: [],
  loadType: 'external',
  defaultUnilateral: false,
  isCustom: false,
  isArchived: false,
  notes: null,
};

describe('ActiveWorkoutScreen (the logging loop)', () => {
  let testDb: TestDb;

  beforeEach(() => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
    useSessionStore.getState().actions.discard();
  });

  afterEach(() => testDb.close());

  it('renders the active session with its exercise and set controls', async () => {
    const { actions } = useSessionStore.getState();
    actions.start(Date.now(), 'Push Day');
    actions.addExercise(bench);

    await render(<ActiveWorkoutScreen />);

    expect(await screen.findByText('Barbell Bench Press')).toBeTruthy();
    expect(screen.getByLabelText('Set 1 weight')).toBeTruthy();
    expect(screen.getByText(/sets done/)).toBeTruthy();
  });

  it('completes a set from the screen (one-tap ✓)', async () => {
    const { actions } = useSessionStore.getState();
    actions.start(Date.now());
    actions.addExercise(bench);
    await render(<ActiveWorkoutScreen />);

    await fireEvent.press(screen.getByLabelText('Complete set 1'));

    await waitFor(() => expect(useSessionStore.getState().exercises[0]?.sets[0]?.done).toBe(true));
  });

  it('shows the empty state before any exercise is added', async () => {
    useSessionStore.getState().actions.start(Date.now());
    await render(<ActiveWorkoutScreen />);

    expect(screen.getByText('Add your first exercise to start logging')).toBeTruthy();
  });
});
