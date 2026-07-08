import { fireEvent, render, screen } from '@testing-library/react-native';

import { Chip } from './Chip';

describe('Chip', () => {
  it('fires onPress and exposes selected state', async () => {
    const onPress = jest.fn();

    await render(<Chip label="Breakfast" selected={false} onPress={onPress} />);
    const chip = screen.getByRole('button', { name: 'Breakfast' });
    await fireEvent.press(chip);

    expect(onPress).toHaveBeenCalledTimes(1);
    expect(chip.props.accessibilityState).toMatchObject({ selected: false });
  });

  it('marks the selected state for accessibility', async () => {
    await render(<Chip label="Lunch" selected onPress={() => undefined} />);

    expect(screen.getByRole('button', { name: 'Lunch' }).props.accessibilityState).toMatchObject({
      selected: true,
    });
  });
});
