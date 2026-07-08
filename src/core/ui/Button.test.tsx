import { fireEvent, render, screen } from '@testing-library/react-native';

import { Button } from './Button';

describe('Button', () => {
  it('renders its label and fires onPress', async () => {
    const onPress = jest.fn();

    await render(<Button label="Save" onPress={onPress} />);
    await fireEvent.press(screen.getByRole('button', { name: 'Save' }));

    expect(onPress).toHaveBeenCalledTimes(1);
  });

  it('does not fire when disabled and exposes the disabled state', async () => {
    const onPress = jest.fn();

    await render(<Button label="Save" onPress={onPress} disabled />);
    const button = screen.getByRole('button', { name: 'Save' });
    await fireEvent.press(button);

    expect(onPress).not.toHaveBeenCalled();
    expect(button.props.accessibilityState).toMatchObject({ disabled: true });
  });

  it('shows a loading indicator and blocks presses while loading', async () => {
    const onPress = jest.fn();

    await render(<Button label="Save" onPress={onPress} loading />);
    await fireEvent.press(screen.getByRole('button', { name: 'Save' }));

    expect(onPress).not.toHaveBeenCalled();
    expect(screen.getByLabelText('Loading')).toBeTruthy();
  });

  it.each(['primary', 'secondary', 'ghost', 'destructive'] as const)(
    'renders the %s variant',
    async (variant) => {
      await render(<Button label={variant} onPress={() => undefined} variant={variant} />);

      expect(screen.getByRole('button', { name: variant })).toBeTruthy();
    },
  );
});
