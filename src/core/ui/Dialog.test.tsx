import { fireEvent, render, screen } from '@testing-library/react-native';

import { triggerHaptic } from '@/core/theme';

import { Dialog } from './Dialog';

jest.mock('@/core/theme', () => {
  const actual = jest.requireActual('@/core/theme');
  return { ...actual, triggerHaptic: jest.fn() };
});

const baseProps = {
  title: 'Delete workout?',
  message: 'This removes 14 sets.',
  confirmLabel: 'Delete',
};

describe('Dialog', () => {
  it('renders title, message, and both actions when visible', async () => {
    await render(
      <Dialog {...baseProps} visible onConfirm={() => undefined} onCancel={() => undefined} />,
    );

    expect(screen.getByText('Delete workout?')).toBeTruthy();
    expect(screen.getByText('This removes 14 sets.')).toBeTruthy();
    expect(screen.getByRole('button', { name: 'Delete' })).toBeTruthy();
    expect(screen.getByRole('button', { name: 'Cancel' })).toBeTruthy();
  });

  it('fires onConfirm with the warning haptic', async () => {
    const onConfirm = jest.fn();

    await render(
      <Dialog {...baseProps} visible onConfirm={onConfirm} onCancel={() => undefined} />,
    );
    await fireEvent.press(screen.getByRole('button', { name: 'Delete' }));

    expect(onConfirm).toHaveBeenCalledTimes(1);
    expect(triggerHaptic).toHaveBeenCalledWith('warning');
  });

  it('fires onCancel from the cancel action', async () => {
    const onCancel = jest.fn();

    await render(<Dialog {...baseProps} visible onConfirm={() => undefined} onCancel={onCancel} />);
    await fireEvent.press(screen.getByRole('button', { name: 'Cancel' }));

    expect(onCancel).toHaveBeenCalledTimes(1);
  });

  it('renders nothing when not visible', async () => {
    await render(
      <Dialog
        {...baseProps}
        visible={false}
        onConfirm={() => undefined}
        onCancel={() => undefined}
      />,
    );

    expect(screen.queryByText('Delete workout?')).toBeNull();
  });
});
