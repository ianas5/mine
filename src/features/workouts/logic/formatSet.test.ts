import { formatSet } from './formatSet';

describe('formatSet', () => {
  it('shows weight × reps for external, trimming whole numbers', () => {
    expect(formatSet(80, 8, 'external')).toBe('80 × 8');
    expect(formatSet(82.5, 6, 'external')).toBe('82.5 × 6');
  });

  it('shows reps only for bodyweight and seconds for timed', () => {
    expect(formatSet(0, 12, 'bodyweight')).toBe('12');
    expect(formatSet(0, 45, 'timed')).toBe('45s');
  });
});
