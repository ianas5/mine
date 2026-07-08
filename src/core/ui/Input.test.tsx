import { fireEvent, render, screen } from '@testing-library/react-native';

import { Input } from './Input';

describe('Input', () => {
  it('renders label and forwards text changes', async () => {
    const onChangeText = jest.fn();

    await render(<Input label="Notes" value="" onChangeText={onChangeText} />);
    await fireEvent.changeText(screen.getByLabelText('Notes'), 'hello');

    expect(onChangeText).toHaveBeenCalledWith('hello');
  });

  it('shows the unit suffix in numeric mode', async () => {
    await render(
      <Input
        label="Weight"
        value="82.4"
        onChangeText={() => undefined}
        keyboardType="decimal-pad"
        unit="KG"
      />,
    );

    expect(screen.getByText('KG')).toBeTruthy();
  });

  it('announces its error as an alert', async () => {
    await render(
      <Input label="Weight" value="" onChangeText={() => undefined} error="Out of range" />,
    );

    expect(screen.getByRole('alert')).toHaveTextContent('Out of range');
  });

  it('calls onBlur when focus leaves', async () => {
    const onBlur = jest.fn();

    await render(<Input label="Notes" value="" onChangeText={() => undefined} onBlur={onBlur} />);
    const field = screen.getByLabelText('Notes');
    await fireEvent(field, 'focus');
    await fireEvent(field, 'blur');

    expect(onBlur).toHaveBeenCalledTimes(1);
  });
});
