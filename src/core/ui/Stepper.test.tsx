import { fireEvent, render, screen } from '@testing-library/react-native';

import { triggerHaptic } from '@/core/theme';

import { Stepper } from './Stepper';

jest.mock('@/core/theme', () => {
  const actual = jest.requireActual('@/core/theme');
  return { ...actual, triggerHaptic: jest.fn() };
});

describe('Stepper', () => {
  it('increments and decrements by the configured step with a haptic tick', async () => {
    const onChange = jest.fn();

    await render(<Stepper value={80} onChange={onChange} step={2.5} />);
    await fireEvent.press(screen.getByRole('button', { name: 'Increase' }));
    await fireEvent.press(screen.getByRole('button', { name: 'Decrease' }));

    expect(onChange).toHaveBeenNthCalledWith(1, 82.5);
    expect(onChange).toHaveBeenNthCalledWith(2, 77.5);
    expect(triggerHaptic).toHaveBeenCalledWith('light');
  });

  it('respects min and max bounds', async () => {
    const onChange = jest.fn();

    await render(<Stepper value={0} onChange={onChange} step={1} min={0} max={1} />);
    await fireEvent.press(screen.getByRole('button', { name: 'Decrease' }));

    expect(onChange).not.toHaveBeenCalled();
  });

  it('avoids float drift on fractional steps', async () => {
    const onChange = jest.fn();

    await render(<Stepper value={0.3} onChange={onChange} step={0.1} />);
    await fireEvent.press(screen.getByRole('button', { name: 'Increase' }));

    expect(onChange).toHaveBeenCalledWith(0.4);
  });

  it('renders the formatted value', async () => {
    await render(<Stepper value={82.5} onChange={() => undefined} format={(v) => `${v} kg`} />);

    expect(screen.getByText('82.5 kg')).toBeTruthy();
  });
});
