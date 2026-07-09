import type { MuscleGroup } from '@/domain/fitness';

import { rangeWindow } from './ranges';
import type { TrainingWorkout } from './trainingData';
import { computeWorkoutAnalytics, type WorkoutAnalyticsInput } from './workoutAnalytics';
import { FIXTURE, FIXTURE_TODAY, sampleTrainingWorkouts } from './trainingFixture';

const baseInput = (over: Partial<WorkoutAnalyticsInput> = {}): WorkoutAnalyticsInput => ({
  workouts: sampleTrainingWorkouts(),
  weighIns: [],
  weeklyWorkoutTarget: 3,
  defaultBodyweightKg: null,
  scheduledWeekdays: [],
  window: rangeWindow('30d', FIXTURE_TODAY),
  today: FIXTURE_TODAY,
  ...over,
});

const groupVolume = (a: ReturnType<typeof computeWorkoutAnalytics>, g: MuscleGroup): number =>
  a.volumeByGroup.find((v) => v.group === g)?.volumeKg ?? 0;

describe('computeWorkoutAnalytics (ANALYTICS §5.1, reconciled to fixture)', () => {
  it('counts countable workouts and this-week progress (never a % mid-week)', () => {
    const a = computeWorkoutAnalytics(baseInput());
    expect(a.totalWorkouts).toBe(4);
    expect(a.consistency.progress).toEqual({ completed: 1, planned: 3 });
    expect(a.consistency.streak).toBe(0); // current week has 1 of 3, prior weeks 1 of 3
    expect(a.frequencyPerWeek).toBeCloseTo(4 / (30 / 7), 4);
  });

  it('sums total volume and per-group volume exactly (primary group only, §3.3)', () => {
    const a = computeWorkoutAnalytics(baseInput());
    expect(a.totalVolumeKg).toBe(FIXTURE.totalVolumeKg);
    expect(groupVolume(a, 'chest')).toBe(FIXTURE.chestVolumeKg);
    expect(groupVolume(a, 'back')).toBe(FIXTURE.backVolumeKg);
    expect(groupVolume(a, 'quads')).toBe(FIXTURE.quadsVolumeKg);
    expect(a.volumeByGroup[0]?.group).toBe('chest'); // ranked by volume desc
  });

  it('computes push:pull and upper:lower balance and flags imbalance (excludes legs/core/other)', () => {
    const a = computeWorkoutAnalytics(baseInput());
    // push (chest 6180) : pull (back 3660) = 1.6885 → outside 0.8–1.25
    expect(a.pushPull.ratio).toBeCloseTo(6180 / 3660, 4);
    expect(a.pushPull.flagged).toBe(true);
    // upper (chest+back 9840) : lower (quads 2100) = 4.686 → outside 1.0–2.0
    expect(a.upperLower.ratio).toBeCloseTo(9840 / 2100, 4);
    expect(a.upperLower.flagged).toBe(true);
  });

  it('most/least trained over 30d includes zero-count groups', () => {
    const a = computeWorkoutAnalytics(baseInput());
    expect(a.mostTrained?.group).toBe('chest');
    expect(a.mostTrained?.workingSets).toBe(12); // 4 sessions × 3 sets
    expect(a.leastTrained?.workingSets).toBe(0); // a never-trained canonical group
  });

  it('key-exercise strength summary trends the most-frequent lifts', () => {
    const a = computeWorkoutAnalytics(baseInput());
    const bench = a.keyExercises[0];
    expect(bench?.name).toBe('Bench Press');
    expect(bench?.sessions).toBe(4);
    expect(bench?.trend.status).toBe('ok');
    if (bench?.trend.status === 'ok') expect(bench.trend.value.classification).toBe('improving');
    // Row (2 sessions) is below the trend minimum → insufficient, never faked.
    const row = a.keyExercises.find((k) => k.name === 'Barbell Row');
    expect(row?.trend.status).toBe('insufficient-data');
  });

  it('averages session duration over workouts with both timestamps only', () => {
    expect(computeWorkoutAnalytics(baseInput()).avgSessionMinutes).toBe(60);
  });

  it('doubles volume for a single-logged unilateral entry (stored marker only, §3.5)', () => {
    const workouts: TrainingWorkout[] = [
      {
        id: 'u1',
        date: '2026-03-09',
        startedAt: null,
        endedAt: null,
        exercises: [
          {
            exerciseId: 'curl',
            name: 'Single-arm Curl',
            primaryMuscleGroup: 'biceps',
            loadType: 'external',
            counting: 'single_doubled',
            sets: [{ weightKg: 20, reps: 10, warmup: false }],
          },
        ],
      },
    ];
    const a = computeWorkoutAnalytics(baseInput({ workouts }));
    expect(groupVolume(a, 'biceps')).toBe(20 * 10 * 2); // doubled from the marker
  });

  it('reports missed workouts only with a weekday schedule, else no-target-set', () => {
    expect(computeWorkoutAnalytics(baseInput()).missedWorkouts).toMatchObject({
      status: 'insufficient-data',
      reason: 'no-target-set',
    });
    // Monday(0)-scheduled: every past Monday in-window has a session, so 0 missed;
    // add a Friday(4) schedule that was never trained → those Fridays count as missed.
    const scheduled = computeWorkoutAnalytics(baseInput({ scheduledWeekdays: [0, 4] }));
    expect(scheduled.missedWorkouts.status).toBe('ok');
    if (scheduled.missedWorkouts.status === 'ok') {
      expect(scheduled.missedWorkouts.value).toBeGreaterThan(0);
    }
  });

  it('compares against the previous equal-length period', () => {
    const a = computeWorkoutAnalytics(baseInput());
    // The 30 days before the window hold no workouts → prev totals are 0, not null.
    expect(a.totalWorkoutsPrev).toBe(0);
    expect(a.totalVolumePrevKg).toBe(0);
  });
});
