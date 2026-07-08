import { fireEvent, render, screen } from '@testing-library/react-native';

import { SegmentedControl } from './SegmentedControl';

const OPTIONS = ['7D', '30D', '3M'] as const;

describe('SegmentedControl', () => {
  it('renders every option as a tab and reports selection changes', async () => {
    const onChange = jest.fn();

    await render(<SegmentedControl options={OPTIONS} selectedIndex={0} onChange={onChange} />);
    for (const option of OPTIONS) {
      expect(screen.getByRole('tab', { name: option })).toBeTruthy();
    }
    await fireEvent.press(screen.getByRole('tab', { name: '30D' }));

    expect(onChange).toHaveBeenCalledWith(1);
  });

  it('exposes the selected tab state', async () => {
    await render(
      <SegmentedControl options={OPTIONS} selectedIndex={2} onChange={() => undefined} />,
    );

    expect(screen.getByRole('tab', { name: '3M' }).props.accessibilityState).toMatchObject({
      selected: true,
    });
    expect(screen.getByRole('tab', { name: '7D' }).props.accessibilityState).toMatchObject({
      selected: false,
    });
  });
});
