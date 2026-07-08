import { render, screen } from '@testing-library/react-native';
import { Text } from 'react-native';

import { Screen } from './Screen';

describe('Screen', () => {
  it('renders children in the plain variant', async () => {
    await render(
      <Screen>
        <Text>content</Text>
      </Screen>,
    );

    expect(screen.getByText('content')).toBeTruthy();
  });

  it('renders children inside a scroll view when scroll is set', async () => {
    await render(
      <Screen scroll>
        <Text>scrollable</Text>
      </Screen>,
    );

    expect(screen.getByText('scrollable')).toBeTruthy();
  });
});
