import { fireEvent, render, screen } from '@testing-library/react-native';

import { ListRow } from './ListRow';

describe('ListRow', () => {
  it('renders title, subtitle, and trailing value', async () => {
    await render(<ListRow title="Bench Press" subtitle="Chest" trailingValue="85 kg" />);

    expect(screen.getByText('Bench Press')).toBeTruthy();
    expect(screen.getByText('Chest')).toBeTruthy();
    expect(screen.getByText('85 kg')).toBeTruthy();
  });

  it('is pressable with an accessibility label when onPress is given', async () => {
    const onPress = jest.fn();

    await render(<ListRow title="Bench Press" onPress={onPress} chevron />);
    await fireEvent.press(screen.getByRole('button', { name: 'Bench Press' }));

    expect(onPress).toHaveBeenCalledTimes(1);
  });

  it('is not pressable without onPress', async () => {
    await render(<ListRow title="Static row" />);

    expect(screen.queryByRole('button')).toBeNull();
  });
});
