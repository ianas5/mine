import { fireEvent, render, screen, waitFor } from '@testing-library/react-native';

import { setDbForTesting } from '@/core/db';
import { createTestDb, type TestDb } from '@/data/testing/createTestDb';
import { seedDatabase } from '@/data/seed/seedDatabase';

import { WorkoutsScreen } from './WorkoutsScreen';

describe('WorkoutsScreen (library, seeded)', () => {
  let testDb: TestDb;

  beforeEach(async () => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
    await seedDatabase();
  });

  afterEach(() => testDb.close());

  it('renders seeded exercises grouped by muscle', async () => {
    await render(<WorkoutsScreen />);

    expect(await screen.findByText('Barbell Bench Press')).toBeTruthy();
    expect(screen.getByText('Chest')).toBeTruthy();
  });

  it('filters the list by the search query', async () => {
    await render(<WorkoutsScreen />);
    await screen.findByText('Barbell Bench Press');

    await fireEvent.changeText(screen.getByLabelText('Search exercises'), 'squat');

    await waitFor(() => expect(screen.getByText('Back Squat')).toBeTruthy());
    expect(screen.queryByText('Barbell Bench Press')).toBeNull();
  });
});
