import { EMPTY_BESTS, type ExerciseBests } from '@/domain/fitness';

import type { SessionSet } from '../stores/useSessionStore';
import { inSessionPrFlags } from './inSessionPrFlags';

let n = 0;
const set = (weightKg: number, reps: number, warmup = false): SessionSet => ({
  localId: `s${(n += 1)}`,
  weightKg,
  reps,
  rpe: null,
  warmup,
  done: false,
});

describe('inSessionPrFlags (optimistic in-session PR marks)', () => {
  it('marks only the running record-setters, not every equal set', () => {
    const baseline: ExerciseBests = { ...EMPTY_BESTS, heaviestWeightKg: 100, bestE1rmKg: 200 };
    const flags = inSessionPrFlags(
      [set(100, 5), set(105, 5), set(105, 5)],
      'external',
      null,
      baseline,
    );
    // 100 ties the baseline (no PR); first 105 is a new best; the second 105 ties it.
    expect(flags).toEqual([false, true, false]);
  });

  it('treats the first working set of a first-ever exercise as a PR', () => {
    const flags = inSessionPrFlags([set(60, 8), set(60, 8)], 'external', null, EMPTY_BESTS);
    expect(flags).toEqual([true, false]);
  });

  it('never marks a warm-up set', () => {
    const flags = inSessionPrFlags([set(200, 3, true)], 'external', null, EMPTY_BESTS);
    expect(flags).toEqual([false]);
  });
});
