import { act, render, screen } from '@testing-library/react-native';

import { triggerHaptic } from '@/core/theme';

import { ToastHost, showToast } from './Toast';

jest.mock('@/core/theme', () => {
  const actual = jest.requireActual('@/core/theme');
  return { ...actual, triggerHaptic: jest.fn() };
});

describe('Toast', () => {
  it('shows a message as an alert and replaces it with the next one (one at a time)', async () => {
    await render(<ToastHost />);

    // Entry opacity animates from 0 under test timers; include hidden elements in queries.
    const hidden = { includeHiddenElements: true } as const;

    await act(async () => showToast('Weight saved'));
    expect(screen.getByText('Weight saved', hidden)).toBeTruthy();

    await act(async () => showToast('Workout saved · 3 PRs'));
    expect(screen.getByText('Workout saved · 3 PRs', hidden)).toBeTruthy();
    expect(screen.queryByText('Weight saved', hidden)).toBeNull();
  });

  it('fires the success haptic for success tone only', async () => {
    await render(<ToastHost />);

    await act(async () => showToast('Neutral'));
    expect(triggerHaptic).not.toHaveBeenCalled();

    await act(async () => showToast('New PR', 'success'));
    expect(triggerHaptic).toHaveBeenCalledWith('success');
  });

  it('renders nothing before any toast is shown', async () => {
    await render(<ToastHost />);

    expect(screen.queryByRole('alert')).toBeNull();
  });
});
