import { fireEvent, render, screen } from '@testing-library/react-native';

import { SetValueControl } from './SetValueControl';

jest.mock('@/core/theme', () => {
  const actual = jest.requireActual('@/core/theme');
  return { ...actual, triggerHaptic: jest.fn() };
});

const setup = async (value: number) => {
  const onChange = jest.fn();
  await render(
    <SetValueControl
      value={value}
      onChange={onChange}
      step={2.5}
      max={1000}
      decimals={1}
      unit="KG"
      accessibilityLabel="weight"
    />,
  );
  return { onChange };
};

describe('SetValueControl', () => {
  it('increments and decrements by the step', async () => {
    const { onChange } = await setup(80);
    await fireEvent.press(screen.getByRole('button', { name: 'Increase weight' }));
    expect(onChange).toHaveBeenCalledWith(82.5);
  });

  it('does not go below zero', async () => {
    const { onChange } = await setup(1);
    await fireEvent.press(screen.getByRole('button', { name: 'Decrease weight' }));
    expect(onChange).toHaveBeenCalledWith(0);
  });

  it('accepts typed input', async () => {
    const { onChange } = await setup(80);
    await fireEvent(screen.getByLabelText('weight'), 'focus');
    await fireEvent.changeText(screen.getByLabelText('weight'), '82.5');
    expect(onChange).toHaveBeenCalledWith(82.5);
  });
});
