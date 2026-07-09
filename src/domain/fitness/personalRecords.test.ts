import {
  computeExerciseBests,
  detectNewPRs,
  type ExerciseSetRow,
  type UnilateralCounting,
} from './index';

interface RowOpts {
  readonly warmup?: boolean;
  readonly counting?: UnilateralCounting;
  readonly order?: number;
}

const row = (
  workoutId: string,
  weightKg: number,
  reps: number,
  opts: RowOpts = {},
): ExerciseSetRow => ({
  workoutId,
  date: '2026-07-01',
  workoutOrder: opts.order ?? 1,
  weightKg,
  reps,
  warmup: opts.warmup ?? false,
  counting: opts.counting ?? 'none',
});

describe('computeExerciseBests (FITNESS_DOMAIN §3.7)', () => {
  it('is all-null with no working sets (a warm-up alone counts for nothing)', () => {
    expect(computeExerciseBests([row('w1', 200, 1, { warmup: true })], 'external', null)).toEqual({
      heaviestWeightKg: null,
      bestE1rmKg: null,
      bestSetVolumeKg: null,
      bestSessionVolumeKg: null,
    });
  });

  it('computes heaviest / best e1RM / best set & session volume', () => {
    const rows = [
      row('w1', 100, 5),
      row('w1', 100, 5), // session w1 volume = 1000
      row('w2', 110, 3, { order: 2 }), // heaviest; session w2 volume = 330
    ];
    const bests = computeExerciseBests(rows, 'external', null);
    expect(bests.heaviestWeightKg).toBe(110);
    expect(bests.bestE1rmKg).toBeCloseTo(110 * 1.1, 4); // 121 > 100×(1+5/30)
    expect(bests.bestSetVolumeKg).toBe(500);
    expect(bests.bestSessionVolumeKg).toBe(1000);
  });

  it('excludes r > 12 sets from the e1RM PR but not from weight (§3.5)', () => {
    const rows = [row('w1', 100, 15), row('w1', 80, 6)];
    const bests = computeExerciseBests(rows, 'external', null);
    expect(bests.heaviestWeightKg).toBe(100);
    expect(bests.bestE1rmKg).toBeCloseTo(80 * (1 + 6 / 30), 4); // only the 6-rep set is eligible
  });

  it('uses single-side load for weight but doubled volume for single_doubled (§3.4)', () => {
    const bests = computeExerciseBests(
      [row('w1', 50, 10, { counting: 'single_doubled' })],
      'external',
      null,
    );
    expect(bests.heaviestWeightKg).toBe(50); // single side
    expect(bests.bestSetVolumeKg).toBe(1000); // 50 × 10 × 2
  });

  it('a timed exercise sets no weight/e1RM/volume records (edge 8)', () => {
    // `timed` reps are seconds; effectiveLoad is 0, so nothing feeds a load record.
    const rows = [row('w1', 0, 60), row('w2', 0, 90, { order: 2 })];
    expect(computeExerciseBests(rows, 'timed', null)).toEqual({
      heaviestWeightKg: null,
      bestE1rmKg: null,
      bestSetVolumeKg: null,
      bestSessionVolumeKg: null,
    });
  });
});

describe('detectNewPRs (strictly greater; §3.7)', () => {
  const prior = [row('w1', 100, 5)];

  it('treats a tie as no PR', () => {
    expect(detectNewPRs(prior, [row('c', 100, 5)], 'external', null)).toEqual([]);
  });

  it('reports every distinct record a heavier top set beats (§3.7 types are independent)', () => {
    // 105×5 beats prior 100×5 on weight, e1RM, set volume (525>500) and session volume.
    const events = detectNewPRs(prior, [row('c', 105, 5)], 'external', null);
    expect(events.map((e) => e.kind).sort()).toEqual([
      'e1rm',
      'sessionVolume',
      'setVolume',
      'weight',
    ]);
    expect(events.find((e) => e.kind === 'weight')?.value).toBe(105);
  });

  it('counts a first-ever performance as PRs (delight registry: incl. first ever)', () => {
    const events = detectNewPRs([], [row('c', 100, 5)], 'external', null);
    expect(events.map((e) => e.kind).sort()).toEqual([
      'e1rm',
      'sessionVolume',
      'setVolume',
      'weight',
    ]);
  });

  it('reports more reps at a load already lifted, but not a brand-new load', () => {
    const repPr = detectNewPRs(prior, [row('c', 100, 6)], 'external', null);
    expect(repPr.find((e) => e.kind === 'repAtLoad')).toEqual({
      kind: 'repAtLoad',
      value: 6,
      loadKg: 100,
    });
    // A never-lifted load sets a weight PR, not a noisy rep PR.
    const newLoad = detectNewPRs(prior, [row('c', 120, 8)], 'external', null);
    expect(newLoad.some((e) => e.kind === 'repAtLoad')).toBe(false);
  });

  it('reports no PRs for a timed exercise, even beating a longer hold (edge 8)', () => {
    const priorTimed = [row('w1', 0, 60)];
    expect(detectNewPRs(priorTimed, [row('c', 0, 120)], 'timed', null)).toEqual([]);
  });

  it('recedes: deleting the record-holding history lowers the bar (recompute)', () => {
    const full = [row('w1', 100, 5), row('w2', 130, 3, { order: 2 })];
    expect(computeExerciseBests(full, 'external', null).heaviestWeightKg).toBe(130);
    // remove w2 → the record recedes to the remaining history
    const afterDelete = full.filter((r) => r.workoutId !== 'w2');
    expect(computeExerciseBests(afterDelete, 'external', null).heaviestWeightKg).toBe(100);
  });
});
