import type { IsoDate } from '@/core/utils';
import type { BodySnapshot } from '@/domain/body';
import type { Phase, PhaseType } from '@/domain/models';
import type { TargetMacros } from '@/domain/nutrition';

import { isOk } from './metricResult';
import type { DailyNutrition } from './nutritionAnalytics';
import { computePhaseReport, type PhaseReportInput } from './phaseAnalytics';
import type { TrainingWorkout } from './trainingData';

const TARGET: TargetMacros = { kcal: 2000, proteinG: 180, carbG: 200, fatG: 60, waterMl: 3000 };

const phase = (over: Partial<Phase> = {}): Phase => ({
  id: 'ph1',
  name: 'Winter Cut',
  type: 'cutting',
  startDate: '2026-01-01',
  endDate: '2026-02-28',
  notes: null,
  createdAt: 1,
  updatedAt: 1,
  ...over,
});

const bench = (id: string, date: IsoDate, weightKg: number): TrainingWorkout => ({
  id,
  date,
  startedAt: 1000,
  endedAt: 2000,
  exercises: [
    {
      exerciseId: 'ex_bench',
      name: 'Bench Press',
      loadType: 'external',
      primaryMuscleGroup: 'chest',
      counting: 'none',
      sets: [{ weightKg, reps: 5, warmup: false }],
    },
  ],
});

const snap = (date: IsoDate, weightKg: number): BodySnapshot =>
  ({
    date,
    weightKg,
    bodyFatPct: null,
    muscleMassKg: null,
    visceralFat: null,
    bmi: null,
    neckCm: null,
    chestCm: null,
    waistCm: null,
    hipsCm: null,
    leftArmCm: null,
    rightArmCm: null,
    leftForearmCm: null,
    rightForearmCm: null,
    leftThighCm: null,
    rightThighCm: null,
    leftCalfCm: null,
    rightCalfCm: null,
  }) as BodySnapshot;

const nutriDay = (date: IsoDate, kcal: number): DailyNutrition => ({
  date,
  totals: { kcal, proteinG: 190, carbG: 180, fatG: 55 },
  target: TARGET,
  waterMl: 3000,
  logged: true,
});

const baseInput = (over: Partial<PhaseReportInput> = {}): PhaseReportInput => ({
  phase: phase(),
  // A prior bench session (before the phase) sets the baseline the in-phase PRs beat.
  workouts: [
    bench('w0', '2025-12-01', 80),
    bench('w1', '2026-01-10', 100),
    bench('w2', '2026-02-05', 100),
  ],
  weighIns: [],
  snapshots: [snap('2026-01-02', 85), snap('2026-02-27', 82)],
  nutritionDays: [],
  weeklyWorkoutTarget: 4,
  defaultBodyweightKg: 82,
  heightCm: 180,
  today: '2026-03-15',
  ...over,
});

describe('computePhaseReport — training summary', () => {
  it('counts only in-phase working workouts and their volume', () => {
    const r = computePhaseReport(baseInput());
    expect(r.training.workouts).toBe(2); // w0 is before the phase
    expect(r.training.totalVolumeKg).toBe(1000); // 100×5 twice, external
    expect(r.training.volumeByGroup).toEqual([{ group: 'chest', volumeKg: 1000, workingSets: 2 }]);
  });

  it('detects PRs set during the phase against prior history', () => {
    const r = computePhaseReport(baseInput());
    expect(r.training.prs).toHaveLength(1);
    expect(r.training.prs[0]!.exerciseId).toBe('ex_bench');
    // 100 kg beats the prior 80 kg on weight, e1RM, set volume and session volume.
    expect(r.training.prs[0]!.kinds).toEqual(
      expect.arrayContaining(['weight', 'e1rm', 'setVolume', 'sessionVolume']),
    );
    expect(r.training.prCount).toBe(r.training.prs[0]!.kinds.length);
  });

  it('reports weekly consistency vs target', () => {
    const c = computePhaseReport(baseInput()).training.avgWeeklyConsistencyPct;
    expect(isOk(c)).toBe(true);
    if (isOk(c)) expect(c.value).toBeGreaterThanOrEqual(0);
  });
});

describe('computePhaseReport — body deltas + rates', () => {
  it('compares first vs last snapshot in the phase', () => {
    const r = computePhaseReport(baseInput());
    expect(isOk(r.bodyDeltas)).toBe(true);
    if (!isOk(r.bodyDeltas)) return;
    const weight = r.bodyDeltas.value.fields.find((f) => f.field === 'weightKg')!;
    expect(weight.deltaAbs).toBe(-3); // 82 − 85
    expect(r.rates.weightDeltaPerWeekKg).toBeLessThan(0);
  });

  it('needs the §5.4 minimums (span ≥ 14 days, ≥ 2 snapshots)', () => {
    const short = computePhaseReport(
      baseInput({ phase: phase({ startDate: '2026-02-20', endDate: '2026-02-28' }) }),
    );
    expect(short.bodyDeltas.status).toBe('insufficient-data');

    const oneSnap = computePhaseReport(baseInput({ snapshots: [snap('2026-01-02', 85)] }));
    expect(oneSnap.bodyDeltas.status).toBe('insufficient-data');
  });
});

describe('computePhaseReport — intent judgment', () => {
  it('reads a cut with weight down as aligned', () => {
    const r = computePhaseReport(baseInput());
    expect(r.nutrition.intent.alignment).toBe('aligned');
  });

  it('flags a lean bulk that lost weight as counter to intent', () => {
    const r = computePhaseReport(baseInput({ phase: phase({ type: 'lean_bulk' }) }));
    expect(r.nutrition.intent.alignment).toBe('counter');
  });

  it('flags a bulk running a calorie deficit even without body data', () => {
    const r = computePhaseReport(
      baseInput({
        phase: phase({ type: 'lean_bulk' }),
        snapshots: [],
        nutritionDays: [
          nutriDay('2026-01-10', 1500),
          nutriDay('2026-01-17', 1500),
          nutriDay('2026-01-24', 1500),
        ],
      }),
    );
    expect(r.nutrition.nutrition.calorieSkew).toBe('under');
    expect(r.nutrition.intent.alignment).toBe('counter');
  });

  it('treats a custom block as unclear (no fixed direction to judge)', () => {
    const r = computePhaseReport(baseInput({ phase: phase({ type: 'custom' as PhaseType }) }));
    expect(r.nutrition.intent.alignment).toBe('unclear');
  });
});

describe('computePhaseReport — phases are context, not prediction', () => {
  it('a completed phase reads identically regardless of today', () => {
    const early = computePhaseReport(baseInput({ today: '2026-03-01' }));
    const late = computePhaseReport(baseInput({ today: '2027-06-01' }));
    expect(late).toEqual(early);
  });

  it('an ongoing phase reports progress to today', () => {
    const r = computePhaseReport(
      baseInput({ phase: phase({ endDate: null }), today: '2026-02-28' }),
    );
    expect(r.ongoing).toBe(true);
    expect(r.window.endDate).toBe('2026-02-28');
  });
});
