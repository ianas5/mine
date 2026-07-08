import { fireEvent, render, screen } from '@testing-library/react-native';
import { Text } from 'react-native';

import { IconButton } from './IconButton';

describe('IconButton', () => {
  it('fires onPress and exposes its required accessibility label', async () => {
    const onPress = jest.fn();

    await render(
      <IconButton icon={<Text>i</Text>} onPress={onPress} accessibilityLabel="Settings" />,
    );
    await fireEvent.press(screen.getByRole('button', { name: 'Settings' }));

    expect(onPress).toHaveBeenCalledTimes(1);
  });

  it('does not fire when disabled', async () => {
    const onPress = jest.fn();

    await render(
      <IconButton icon={<Text>i</Text>} onPress={onPress} accessibilityLabel="Settings" disabled />,
    );
    await fireEvent.press(screen.getByRole('button', { name: 'Settings' }));

    expect(onPress).not.toHaveBeenCalled();
  });
});
