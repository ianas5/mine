import { fireEvent, render, screen } from '@testing-library/react-native';
import { Text } from 'react-native';

import { Section } from './Section';

describe('Section', () => {
  it('renders its title and children', async () => {
    await render(
      <Section title="This Week">
        <Text>content</Text>
      </Section>,
    );

    expect(screen.getByText('This Week')).toBeTruthy();
    expect(screen.getByText('content')).toBeTruthy();
  });

  it('fires the optional action', async () => {
    const onPress = jest.fn();

    await render(
      <Section title="Workouts" action={{ label: 'See all', onPress }}>
        <Text>content</Text>
      </Section>,
    );
    await fireEvent.press(screen.getByRole('button', { name: 'See all' }));

    expect(onPress).toHaveBeenCalledTimes(1);
  });
});
