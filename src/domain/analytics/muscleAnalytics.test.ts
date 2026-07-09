import { computeMuscleReports, reportForGroup, type MuscleAnalyticsInput } from './muscleAnalytics';
import { rangeWindow } from './ranges';
import { FIXTURE_TODAY, sampleTrainingWorkouts } from './trainingFixture';

const input: MuscleAnalyticsInput = {
  workouts: sampleTrainingWorkouts(),
  weighIns: [],
  defaultBodyweightKg: null,
  window: rangeWindow('90d', FIXTURE_TODAY),
  today: FIXTURE_TODAY,
};

describe('reportForGroup (ANALYTICS §5.6)', () => {
  it('reviews a trained group like a coach (strongest, improving, last trained)', () => {
    const chest = reportForGroup(input, 'chest');
    expect(chest.untrained).toBe(false);
    expect(chest.volume30dKg).toBe(6180);
    expect(chest.workingSets30d).toBe(12);
    expect(chest.currentWeekVolumeKg).toBe(1590); // only the 2026-03-09 session
    expect(chest.strongest?.name).toBe('Bench Press');
    expect(chest.strongest?.value).toBeCloseTo(106 * (1 + 5 / 30), 3); // best e1RM
    expect(chest.fastestImproving?.name).toBe('Bench Press');
    expect(chest.fastestImproving?.value).toBeGreaterThan(0); // rising e1RM slope
    expect(chest.lastTrained).toEqual({ date: '2026-03-09', daysSince: 6 });
  });

  it('gives an honest zero state for a never-trained group', () => {
    const shoulders = reportForGroup(input, 'shoulders');
    expect(shoulders.untrained).toBe(true);
    expect(shoulders.volume30dKg).toBe(0);
    expect(shoulders.workingSets30d).toBe(0);
    expect(shoulders.strongest).toBeNull();
    expect(shoulders.fastestImproving).toBeNull();
    expect(shoulders.lastTrained).toBeNull();
    expect(shoulders.frequencyPerWeek).toBe(0);
  });
});

describe('computeMuscleReports', () => {
  it('reports every canonical group (11, excluding Other)', () => {
    const reports = computeMuscleReports(input);
    expect(reports).toHaveLength(11);
    expect(reports.map((r) => r.group)).not.toContain('other');
    expect(
      reports
        .filter((r) => !r.untrained)
        .map((r) => r.group)
        .sort(),
    ).toEqual(['back', 'chest', 'quads']);
  });
});
