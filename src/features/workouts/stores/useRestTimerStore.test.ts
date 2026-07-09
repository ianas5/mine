import { formatCountdown } from '../hooks/useRestCountdown';
import { DEFAULT_REST_SEC, REST_EXTEND_SEC, useRestTimerStore } from './useRestTimerStore';

const s = () => useRestTimerStore.getState();

describe('useRestTimerStore', () => {
  beforeEach(() => s().actions.reset());

  it('auto-starts a wall-clock rest at the default duration', () => {
    s().actions.start('ex1', 1000);
    expect(s().running).toBe(true);
    expect(s().durationSec).toBe(DEFAULT_REST_SEC);
    expect(s().endsAt).toBe(1000 + DEFAULT_REST_SEC * 1000);
  });

  it('extends the running rest and remembers it for that exercise', () => {
    s().actions.start('ex1', 1000);
    s().actions.extend(REST_EXTEND_SEC, 1000);
    expect(s().durationSec).toBe(DEFAULT_REST_SEC + REST_EXTEND_SEC);

    // The next set of the same exercise starts from the remembered, longer rest.
    s().actions.skip();
    s().actions.start('ex1', 5000);
    expect(s().durationSec).toBe(DEFAULT_REST_SEC + REST_EXTEND_SEC);
  });

  it('skip clears the countdown but keeps prefs; reset forgets everything', () => {
    s().actions.start('ex1', 1000);
    s().actions.extend(REST_EXTEND_SEC, 1000);
    s().actions.skip();
    expect(s().running).toBe(false);
    expect(s().endsAt).toBeNull();

    s().actions.reset();
    s().actions.start('ex1', 2000);
    expect(s().durationSec).toBe(DEFAULT_REST_SEC);
  });
});

describe('formatCountdown', () => {
  it('renders m:ss, rounding up remaining time', () => {
    expect(formatCountdown(90_000)).toBe('1:30');
    expect(formatCountdown(1_500)).toBe('0:02');
    expect(formatCountdown(0)).toBe('0:00');
  });
});
