import { formatRelativeDate, isoWeekday, todayIso, weekdayLabel } from './date';

describe('date helpers (FITNESS_DOMAIN §2.3, edge 14 device-local)', () => {
  it('formats today as device-local YYYY-MM-DD (never UTC-shifted)', () => {
    // Local components, late in the day — a naive UTC conversion would roll the date.
    expect(todayIso(new Date(2026, 6, 9, 23, 30))).toBe('2026-07-09');
    expect(todayIso(new Date(2026, 0, 5, 0, 15))).toBe('2026-01-05');
  });

  it('maps ISO weekday with Monday = 0 … Sunday = 6', () => {
    // 2024-01-01 was a Monday; the week runs Mon→Sun.
    expect(isoWeekday('2024-01-01')).toBe(0); // Monday
    expect(isoWeekday('2024-01-02')).toBe(1); // Tuesday
    expect(isoWeekday('2024-01-03')).toBe(2); // Wednesday
    expect(isoWeekday('2024-01-06')).toBe(5); // Saturday
    expect(isoWeekday('2024-01-07')).toBe(6); // Sunday
  });

  it('labels ISO weekdays consistently with the conversion', () => {
    expect(weekdayLabel(0)).toBe('Monday');
    expect(weekdayLabel(6)).toBe('Sunday');
    expect(weekdayLabel(isoWeekday('2024-01-07'))).toBe('Sunday');
  });

  it('humanizes recency relative to a device-local now', () => {
    const now = new Date(2026, 6, 9); // Thu 2026-07-09
    expect(formatRelativeDate('2026-07-09', now)).toBe('Today');
    expect(formatRelativeDate('2026-07-08', now)).toBe('Yesterday');
    expect(formatRelativeDate('2026-07-05', now)).toBe('Sun · 4d ago');
  });
});
