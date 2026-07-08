import {
  effectiveLoadKg,
  epleyOneRepMax,
  isE1rmEligible,
  isVolumeLowConfidence,
  isWorkingSet,
  setVolumeKg,
} from './sets';

const set = (weightKg: number, reps: number, warmup = false) => ({ weightKg, reps, warmup });

describe('effectiveLoadKg (FITNESS_DOMAIN §3.4)', () => {
  it('returns the logged weight for external load', () => {
    expect(effectiveLoadKg('external', 80, null)).toBe(80);
  });
  it('uses bodyweight for bodyweight, 0 when unknown', () => {
    expect(effectiveLoadKg('bodyweight', 0, 82)).toBe(82);
    expect(effectiveLoadKg('bodyweight', 0, null)).toBe(0);
  });
  it('adds load for bodyweight_plus and subtracts (floored) for assisted', () => {
    expect(effectiveLoadKg('bodyweight_plus', 20, 82)).toBe(102);
    expect(effectiveLoadKg('assisted', 30, 82)).toBe(52);
    expect(effectiveLoadKg('assisted', 100, 82)).toBe(0);
  });
  it('is 0 for timed', () => {
    expect(effectiveLoadKg('timed', 0, 82)).toBe(0);
  });
});

describe('isWorkingSet (FITNESS_DOMAIN §3.2)', () => {
  it('excludes warm-ups (edge 1)', () => {
    expect(isWorkingSet(set(80, 8, true), 'external')).toBe(false);
  });
  it('excludes reps < 1', () => {
    expect(isWorkingSet(set(80, 0), 'external')).toBe(false);
  });
  it('requires positive weight for external', () => {
    expect(isWorkingSet(set(0, 8), 'external')).toBe(false);
    expect(isWorkingSet(set(80, 8), 'external')).toBe(true);
  });
  it('counts bodyweight/timed stimulus even without weight (edge 5, 8)', () => {
    expect(isWorkingSet(set(0, 8), 'bodyweight')).toBe(true);
    expect(isWorkingSet(set(0, 45), 'timed')).toBe(true);
  });
});

describe('setVolumeKg (FITNESS_DOMAIN §3.5)', () => {
  it('is weight × reps for external working sets', () => {
    expect(setVolumeKg(set(80, 8), 'external', null, 'none')).toBe(640);
  });
  it('doubles for a single-logged unilateral entry (edge 6)', () => {
    expect(setVolumeKg(set(20, 10), 'external', null, 'single_doubled')).toBe(400);
    expect(setVolumeKg(set(20, 10), 'external', null, 'per_side')).toBe(200);
  });
  it('is 0 for warm-ups and timed (edge 1, 8)', () => {
    expect(setVolumeKg(set(80, 8, true), 'external', null, 'none')).toBe(0);
    expect(setVolumeKg(set(0, 45), 'timed', 82, 'none')).toBe(0);
  });
});

describe('isVolumeLowConfidence (edge 5)', () => {
  it('flags a bodyweight working set with unknown bodyweight', () => {
    expect(isVolumeLowConfidence(set(0, 8), 'bodyweight', null)).toBe(true);
  });
  it('does not flag when bodyweight is known or load is external', () => {
    expect(isVolumeLowConfidence(set(0, 8), 'bodyweight', 82)).toBe(false);
    expect(isVolumeLowConfidence(set(80, 8), 'external', null)).toBe(false);
  });
});

describe('epleyOneRepMax + eligibility (FITNESS_DOMAIN §3.5)', () => {
  it('applies the Epley formula', () => {
    expect(epleyOneRepMax(100, 1)).toBe(100);
    expect(epleyOneRepMax(100, 10)).toBeCloseTo(133.33, 1);
    expect(epleyOneRepMax(0, 5)).toBe(0);
  });
  it('trusts e1RM only at reps 1..12 (edge 7)', () => {
    expect(isE1rmEligible(12)).toBe(true);
    expect(isE1rmEligible(13)).toBe(false);
    expect(isE1rmEligible(0)).toBe(false);
  });
});
