import { render, screen } from '@testing-library/react-native';
import { Text } from 'react-native';

import { Card } from './Card';

describe('Card', () => {
  it.each(['default', 'raised', 'accentEdge'] as const)(
    'renders children in the %s variant',
    async (variant) => {
      await render(
        <Card variant={variant}>
          <Text>{variant}</Text>
        </Card>,
      );

      expect(screen.getByText(variant)).toBeTruthy();
    },
  );
});
