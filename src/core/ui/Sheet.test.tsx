import { fireEvent, render, screen } from '@testing-library/react-native';
import { Text } from 'react-native';

import { Sheet } from './Sheet';

describe('Sheet', () => {
  it('renders title and children when visible', async () => {
    await render(
      <Sheet visible onClose={() => undefined} title="Add Weight">
        <Text>fields</Text>
      </Sheet>,
    );

    expect(screen.getByText('Add Weight')).toBeTruthy();
    expect(screen.getByText('fields')).toBeTruthy();
  });

  it('closes directly when clean', async () => {
    const onClose = jest.fn();

    await render(
      <Sheet visible onClose={onClose}>
        <Text>fields</Text>
      </Sheet>,
    );
    await fireEvent.press(screen.getByLabelText('Dismiss sheet'));

    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it('guards dismissal with a discard dialog when dirty', async () => {
    const onClose = jest.fn();

    await render(
      <Sheet visible onClose={onClose} dirty>
        <Text>fields</Text>
      </Sheet>,
    );
    await fireEvent.press(screen.getByLabelText('Dismiss sheet'));

    expect(onClose).not.toHaveBeenCalled();
    expect(screen.getByText('Discard entry?')).toBeTruthy();

    await fireEvent.press(screen.getByRole('button', { name: 'Discard' }));
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it('keeps the sheet open when discard is cancelled', async () => {
    const onClose = jest.fn();

    await render(
      <Sheet visible onClose={onClose} dirty>
        <Text>fields</Text>
      </Sheet>,
    );
    await fireEvent.press(screen.getByLabelText('Dismiss sheet'));
    await fireEvent.press(screen.getByRole('button', { name: 'Cancel' }));

    expect(onClose).not.toHaveBeenCalled();
    expect(screen.getByText('fields')).toBeTruthy();
  });
});
