import { addDaysIso, daysBetweenIso, type IsoDate } from '@/core/utils';
import { weekStartIso } from '@/domain/fitness';
import { dayAdherence, type MacroSet, type TargetMacros } from '@/domain/nutrition';

import { insufficient, ok, type MetricResult, type Range } from './metricResult';
import { computeTrend, type Trend } from './trend';
import type { SeriesPoint } from './timeSeries';

/** One day's nutrition, target resolved by the repository (never here — DATABASE rule 13). */
export interface DailyNutrition {
  readonly date: IsoDate;
  readonly totals: MacroSet;
  readonly target: TargetMacros | null;
  readonly waterMl: number | null;
  /** Had ≥ 1 meal entry — the honest definition of a "logged" nutrition day (§4.2). */
  readonly logged: boolean;
}

export interface AdherenceStat {
  readonly pct: number;
  readonly hitDays: number;
  /** Logged days that had a target to judge against. */
  readonly loggedDays: number;
}

export interface NutritionAnalytics {
  readonly daysInRange: number;
  readonly loggedDays: number;
  /** logged ÷ days in range (§5.2 logging completeness). */
  readonly completeness: number;
  /** Averages over logged days, or null when nothing is logged. */
  readonly avg: MacroSet | null;
  readonly activeTarget: TargetMacros | null;
  readonly calorieAdherence: MetricResult<AdherenceStat>;
  readonly proteinAdherence: MetricResult<AdherenceStat>;
  readonly carbAdherence: MetricResult<AdherenceStat>;
  readonly fatAdherence: MetricResult<AdherenceStat>;
  readonly waterAdherence: MetricResult<AdherenceStat>;
  /** Which side calorie misses skew to (§5.2), or null when there are no misses. */
  readonly calorieSkew: 'under' | 'over' | null;
  readonly proteinTrend: MetricResult<Trend>;
}

export interface NutritionAnalyticsInput {
  readonly days: readonly DailyNutrition[];
  readonly window: Range;
  readonly today: IsoDate;
}

const inWindow = (days: readonly DailyNutrition[], window: Range): DailyNutrition[] =>
  days.filter(
    (d) => (window.startDate === null || d.date >= window.startDate) && d.date <= window.endDate,
  );

const loggedWithTarget = (days: readonly DailyNutrition[]): DailyNutrition[] =>
  days.filter((d) => d.logged && d.target !== null);

/** Whether a day hit a given macro (FITNESS_DOMAIN §4.3). */
function hit(
  day: DailyNutrition,
  macro: 'calories' | 'protein' | 'carbs' | 'fat' | 'water',
): boolean {
  if (day.target === null) return false;
  const a = dayAdherence(day.target, day.totals, day.waterMl ?? 0);
  if (macro === 'water') return a.water === 'hit';
  return a[macro] === 'hit';
}

function adherence(
  days: readonly DailyNutrition[],
  macro: 'calories' | 'protein' | 'carbs' | 'fat',
): MetricResult<AdherenceStat> {
  const judged = loggedWithTarget(days);
  if (judged.length === 0) {
    return insufficient('no-data', 'Log a few days against a target to see adherence');
  }
  const hitDays = judged.filter((d) => hit(d, macro)).length;
  return ok(
    { pct: Math.round((hitDays / judged.length) * 100), hitDays, loggedDays: judged.length },
    {
      key: '30d',
      startDate: judged[0]!.date,
      endDate: judged[judged.length - 1]!.date,
      days: null,
    },
    { points: judged.length, spanDays: 0 },
  );
}

/** Consecutive **logged** days below the protein target, ending at the latest logged day
 * (§4.3). Unlogged days are skipped (no data ≠ a miss); the run is over logged days. */
export function proteinMissStreak(days: readonly DailyNutrition[]): number {
  const logged = loggedWithTarget(days)
    .slice()
    .sort((a, b) => (a.date < b.date ? 1 : -1)); // newest first
  let streak = 0;
  for (const day of logged) {
    if (hit(day, 'protein')) break;
    streak += 1;
  }
  return streak;
}

/** Signals over the trailing `n` days from `today` (fixed window, for insight rules). */
export function trailingSignals(
  days: readonly DailyNutrition[],
  today: IsoDate,
  n: number,
): { loggedCount: number; proteinHits: number; waterHits: number; waterTargetDays: number } {
  const start = addDaysIso(today, -(n - 1));
  const recent = days.filter((d) => d.date >= start && d.date <= today);
  const logged = recent.filter((d) => d.logged);
  return {
    loggedCount: logged.length,
    proteinHits: logged.filter((d) => d.target !== null && hit(d, 'protein')).length,
    waterHits: logged.filter((d) => d.target?.waterMl != null && hit(d, 'water')).length,
    waterTargetDays: logged.filter((d) => d.target?.waterMl != null).length,
  };
}

/** Calorie-miss skew over the trailing `n` days: 'under'/'over' when ≥ 70% of misses
 * fall one side, else null (insight rule 10 uses n = 14). */
export function calorieSkew(
  days: readonly DailyNutrition[],
  today: IsoDate,
  n: number,
): 'under' | 'over' | null {
  const start = addDaysIso(today, -(n - 1));
  const judged = loggedWithTarget(days.filter((d) => d.date >= start && d.date <= today));
  let under = 0;
  let over = 0;
  for (const day of judged) {
    const status = dayAdherence(day.target!, day.totals, day.waterMl ?? 0).calories;
    if (status === 'under') under += 1;
    else if (status === 'over') over += 1;
  }
  const misses = under + over;
  if (misses === 0) return null;
  if (under / misses >= 0.7) return 'under';
  if (over / misses >= 0.7) return 'over';
  return null;
}

function meanMacros(days: readonly DailyNutrition[]): MacroSet | null {
  const logged = days.filter((d) => d.logged);
  if (logged.length === 0) return null;
  const sum = logged.reduce(
    (acc, d) => ({
      kcal: acc.kcal + d.totals.kcal,
      proteinG: acc.proteinG + d.totals.proteinG,
      carbG: acc.carbG + d.totals.carbG,
      fatG: acc.fatG + d.totals.fatG,
    }),
    { kcal: 0, proteinG: 0, carbG: 0, fatG: 0 },
  );
  return {
    kcal: sum.kcal / logged.length,
    proteinG: sum.proteinG / logged.length,
    carbG: sum.carbG / logged.length,
    fatG: sum.fatG / logged.length,
  };
}

/** Weekly protein trend over the window: mean protein of weeks with ≥ 4 logged days (§5.2). */
function proteinTrend(days: readonly DailyNutrition[], window: Range): MetricResult<Trend> {
  const byWeek = new Map<IsoDate, number[]>();
  for (const day of days) {
    if (!day.logged) continue;
    const week = weekStartIso(day.date);
    const bucket = byWeek.get(week);
    if (bucket) bucket.push(day.totals.proteinG);
    else byWeek.set(week, [day.totals.proteinG]);
  }
  const series: SeriesPoint[] = [...byWeek.entries()]
    .filter(([, values]) => values.length >= 4)
    .map(([date, values]) => ({ date, value: values.reduce((s, v) => s + v, 0) / values.length }))
    .sort((a, b) => (a.date < b.date ? -1 : 1));
  return computeTrend(
    series,
    { stabilityThreshold: 8, goodDirection: 'higher', pointNoun: 'full weeks of logging' },
    window,
  );
}

/**
 * The Nutrition analytics calculator (ANALYTICS §5.2), pure. Everything is over **logged
 * days only** (unlogged ≠ failed, ≠ succeeded — §2/§4.3); logging completeness is its own
 * honest number. Adherence uses the §4.3 rules against each day's resolved target.
 */
export function computeNutritionAnalytics(input: NutritionAnalyticsInput): NutritionAnalytics {
  const windowed = inWindow(input.days, input.window);
  const loggedDays = windowed.filter((d) => d.logged).length;
  const daysInRange =
    input.window.days ??
    Math.max(1, daysBetweenIso(windowed[0]?.date ?? input.today, input.today) + 1);

  const waterJudged = loggedWithTarget(windowed).filter((d) => d.target?.waterMl != null);
  const waterAdherence: MetricResult<AdherenceStat> =
    waterJudged.length === 0
      ? insufficient('no-target-set', 'Set a water goal to track water adherence')
      : ok(
          {
            pct: Math.round(
              (waterJudged.filter((d) => hit(d, 'water')).length / waterJudged.length) * 100,
            ),
            hitDays: waterJudged.filter((d) => hit(d, 'water')).length,
            loggedDays: waterJudged.length,
          },
          input.window,
          { points: waterJudged.length, spanDays: 0 },
        );

  const activeTarget = [...windowed].reverse().find((d) => d.target !== null)?.target ?? null;

  return {
    daysInRange,
    loggedDays,
    completeness: daysInRange > 0 ? loggedDays / daysInRange : 0,
    avg: meanMacros(windowed),
    activeTarget,
    calorieAdherence: adherence(windowed, 'calories'),
    proteinAdherence: adherence(windowed, 'protein'),
    carbAdherence: adherence(windowed, 'carbs'),
    fatAdherence: adherence(windowed, 'fat'),
    waterAdherence,
    calorieSkew: calorieSkew(windowed, input.today, Math.max(1, daysInRange)),
    proteinTrend: proteinTrend(windowed, input.window),
  };
}
