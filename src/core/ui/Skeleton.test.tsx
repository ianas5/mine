import { render, screen } from '@testing-library/react-native';

import { Skeleton } from './Skeleton';

describe('Skeleton', () => {
  it('renders with a loading accessibility label', async () => {
    await render(<Skeleton />);

    expect(screen.getByLabelText('Loading')).toBeTruthy();
  });

  it('accepts explicit dimensions', async () => {
    await render(<Skeleton width={120} height={32} radius={8} />);

    const node = screen.getByLabelText('Loading');
    expect(node.props.style).toMatchObject({ width: 120, height: 32, borderRadius: 8 });
  });
});
