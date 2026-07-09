import { defaultSlotForHour } from './taxonomy';

describe('defaultSlotForHour (UI_UX §5.2 time-of-day fallback)', () => {
  it('maps the hour to a plausible meal slot', () => {
    expect(defaultSlotForHour(7)).toBe('breakfast');
    expect(defaultSlotForHour(12)).toBe('lunch');
    expect(defaultSlotForHour(19)).toBe('dinner');
    expect(defaultSlotForHour(23)).toBe('snacks');
  });

  it('uses the boundaries consistently', () => {
    expect(defaultSlotForHour(11)).toBe('lunch');
    expect(defaultSlotForHour(15)).toBe('dinner');
    expect(defaultSlotForHour(21)).toBe('snacks');
  });
});
