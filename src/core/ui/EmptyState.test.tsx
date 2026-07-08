import { fireEvent, render, screen } from '@testing-library/react-native';

import { EmptyState } from './EmptyState';

describe('EmptyState', () => {
  it('renders the factual line', async () => {
    await render(<EmptyState title="No weigh-ins yet" />);

    expect(screen.getByText('No weigh-ins yet')).toBeTruthy();
  });

  it('renders and fires the optional CTA', async () => {
    const onPress = jest.fn();

    await render(<EmptyState title="No weigh-ins yet" cta={{ label: 'Add Weight', onPress }} />);
    await fireEvent.press(screen.getByRole('button', { name: 'Add Weight' }));

    expect(onPress).toHaveBeenCalledTimes(1);
  });
});
