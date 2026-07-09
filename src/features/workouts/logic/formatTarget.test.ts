import { formatTarget } from './formatTarget';

describe('formatTarget', () => {
  it('renders sets × rep-range with RPE', () => {
    expect(formatTarget({ sets: 3, repMin: 8, repMax: 10, rpe: 8 })).toBe(
      'Target · 3 × 8–10 @ RPE 8',
    );
  });

  it('collapses an equal rep range and omits missing pieces', () => {
    expect(formatTarget({ sets: 5, repMin: 5, repMax: 5, rpe: null })).toBe('Target · 5 × 5');
    expect(formatTarget({ sets: null, repMin: 8, repMax: null, rpe: null })).toBe(
      'Target · 8+ reps',
    );
    expect(formatTarget({ sets: 4, repMin: null, repMax: null, rpe: null })).toBe(
      'Target · 4 sets',
    );
  });

  it('returns null when the template set no targets (never fabricate, P8)', () => {
    expect(formatTarget(null)).toBeNull();
    expect(formatTarget({ sets: null, repMin: null, repMax: null, rpe: null })).toBeNull();
  });
});
