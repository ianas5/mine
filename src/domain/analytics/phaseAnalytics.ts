import { addDaysIso, daysBetweenIso, type IsoDate } from '@/core/utils';
import { compareSnapshots, type BodySnapshot, type FieldComparison } from '@/domain/body';
import {
  countableWorkoutsByWeek,
  detectNewPRs,
  isWorkingSet,
  weekConsistencyPercent,
  weekStartIso,
  type ExerciseSetRow,
  type PrKind,
} from '@/domain/fitness';
import { PHASE_INTENT, PHASE_TYPE_LABELS, type Phase, type PhaseType } from '@/domain/models';

import { insufficient, ok, isOk, type MetricResult, type Range } from './metricResult';
import {
  computeNutritionAnalytics,
  type DailyNutrition,
  type NutritionAnalytics,
} from './nutritionAnalytics';
import { resolveBodyweightForDate, type TrainingWorkout, type WeighIn } from './trainingData';
import {
  computeWorkoutAnalytics,
  type GroupVolume,
  type KeyExerciseStrength,
} from './workoutAnalytics';

/** §5.4/§3.1 minimums for a body-delta comparison: a real span and two points. */
export const PHASE_MIN_DAYS = 14;
export const PHASE_MIN_SNAPSHOTS = 2;
/** Weight-stability deadband (kg) — matches FITNESS_DOMAIN §6.4 / BODY_STABILITY.weightKg. */
const WEIGHT_STABLE_KG = 0.8;
/** A "steady" block that drifts more than this (kg) over its length ran counter to intent. */
const STABLE_DRIFT_KG = 2;

export interface PhaseBodyDeltas {
  readonly firstDate: IsoDate;
  readonly lastDate: IsoDate;
  /** Per-field first-vs-last comparison within the phase (FITNESS_DOMAIN §5.4). */
  readonly fields: readonly FieldComparison[];
}

export interface PhasePr {
  readonly exerciseId: string;
  readonly name: string;
  readonly kinds: readonly PrKind[];
}

export interface PhaseTrainingSummary {
  readonly workouts: number;
  readonly weeks: number;
  readonly workoutsPerWeek: number;
  /** Mean weekly consistency vs. target across the phase weeks. */
  readonly avgWeeklyConsistencyPct: MetricResult<number>;
  readonly totalVolumeKg: number;
  readonly volumeByGroup: readonly GroupVolume[];
  /** Top lifts with their e1RM change over the block (§5.4 key-exercise strength). */
  readonly keyExercises: readonly KeyExerciseStrength[];
  readonly prCount: number;
  readonly prs: readonly PhasePr[];
}

export type IntentAlignment = 'aligned' | 'counter' | 'unclear';

export interface PhaseIntentVerdict {
  readonly alignment: IntentAlignment;
  /** One honest sentence judging the block against its declared intent (§5.4). */
  readonly message: string;
}

export interface PhaseNutritionSummary {
  readonly nutrition: NutritionAnalytics;
  readonly intent: PhaseIntentVerdict;
}

/** Per-week normalization so blocks of different lengths compare fairly (§5.4). */
export interface PhaseWeeklyRates {
  readonly weeks: number;
  readonly workoutsPerWeek: number;
  readonly volumePerWeekKg: number;
  /** (last − first) body weight ÷ weeks, or null when body deltas are insufficient. */
  readonly weightDeltaPerWeekKg: number | null;
}

export interface PhaseReport {
  readonly phase: Phase;
  readonly ongoing: boolean;
  readonly window: Range;
  readonly spanDays: number;
  readonly bodyDeltas: MetricResult<PhaseBodyDeltas>;
  readonly training: PhaseTrainingSummary;
  readonly nutrition: PhaseNutritionSummary;
  readonly rates: PhaseWeeklyRates;
}

export interface PhaseReportInput {
  readonly phase: Phase;
  readonly workouts: readonly TrainingWorkout[];
  readonly weighIns: readonly WeighIn[];
  readonly snapshots: readonly BodySnapshot[];
  readonly nutritionDays: readonly DailyNutrition[];
  readonly weeklyWorkoutTarget: number;
  readonly defaultBodyweightKg: number | null;
  readonly heightCm: number | null;
  readonly today: IsoDate;
}

const round1 = (n: number): number => Math.round(n * 10) / 10;
const signed = (n: number): string => `${n > 0 ? '+' : ''}${round1(n)}`;
const inRange = (date: IsoDate, start: IsoDate, end: IsoDate): boolean =>
  date >= start && date <= end;

/** First-vs-last body comparison across the phase, gated by the §5.4 minimums. */
function bodyDeltas(
  snapshots: readonly BodySnapshot[],
  start: IsoDate,
  end: IsoDate,
  spanDays: number,
  heightCm: number | null,
): MetricResult<PhaseBodyDeltas> {
  const within = snapshots
    .filter((s) => inRange(s.date, start, end))
    .slice()
    .sort((a, b) => (a.date < b.date ? -1 : 1));
  if (spanDays < PHASE_MIN_DAYS) {
    return insufficient('span-too-short', `A phase needs ${PHASE_MIN_DAYS} days for body deltas`);
  }
  if (within.length < PHASE_MIN_SNAPSHOTS) {
    return insufficient(
      'too-few-points',
      `Log at least ${PHASE_MIN_SNAPSHOTS} measurements in the phase`,
    );
  }
  const first = within[0]!;
  const last = within[within.length - 1]!;
  return ok(
    { firstDate: first.date, lastDate: last.date, fields: compareSnapshots(first, last, heightCm) },
    { key: 'all', startDate: start, endDate: end, days: spanDays },
    { points: within.length, spanDays: daysBetweenIso(first.date, last.date) },
  );
}

/** Mean weekly consistency (vs. `target`) over every week the phase spans. */
function avgWeeklyConsistency(
  countableDates: readonly IsoDate[],
  start: IsoDate,
  end: IsoDate,
  target: number,
): MetricResult<number> {
  const byWeek = countableWorkoutsByWeek(countableDates);
  const weekStarts: IsoDate[] = [];
  for (let w = weekStartIso(start); w <= end; w = addDaysIso(w, 7)) {
    weekStarts.push(w);
  }
  if (weekStarts.length === 0) {
    return insufficient('span-too-short', 'A phase needs at least a full week for consistency');
  }
  const total = weekStarts.reduce(
    (sum, ws) => sum + weekConsistencyPercent(byWeek.get(ws) ?? 0, target),
    0,
  );
  return ok(
    Math.round(total / weekStarts.length),
    { key: 'all', startDate: start, endDate: end, days: null },
    {
      points: weekStarts.length,
      spanDays: daysBetweenIso(start, end),
    },
  );
}

/** PRs set during the phase: in-phase sets judged against all prior history per lift. */
function phasePrs(
  workouts: readonly TrainingWorkout[],
  start: IsoDate,
  end: IsoDate,
  bodyweight: number | null,
): { prs: PhasePr[]; prCount: number } {
  const byExercise = new Map<
    string,
    {
      name: string;
      loadType: TrainingWorkout['exercises'][number]['loadType'];
      prior: ExerciseSetRow[];
      inPhase: ExerciseSetRow[];
    }
  >();
  for (const w of workouts) {
    if (w.date > end) continue;
    for (const ex of w.exercises) {
      let acc = byExercise.get(ex.exerciseId);
      if (!acc) {
        acc = { name: ex.name, loadType: ex.loadType, prior: [], inPhase: [] };
        byExercise.set(ex.exerciseId, acc);
      }
      for (const s of ex.sets) {
        const row: ExerciseSetRow = {
          workoutId: w.id,
          date: w.date,
          workoutOrder: w.startedAt ?? 0,
          weightKg: s.weightKg,
          reps: s.reps,
          warmup: s.warmup,
          counting: ex.counting,
        };
        if (w.date < start) acc.prior.push(row);
        else acc.inPhase.push(row);
      }
    }
  }
  const prs: PhasePr[] = [];
  for (const [exerciseId, acc] of byExercise) {
    if (acc.inPhase.length === 0) continue;
    const events = detectNewPRs(acc.prior, acc.inPhase, acc.loadType, bodyweight);
    if (events.length > 0) {
      prs.push({ exerciseId, name: acc.name, kinds: events.map((e) => e.kind) });
    }
  }
  return { prs, prCount: prs.reduce((n, p) => n + p.kinds.length, 0) };
}

/** Judges the block against its declared intent (§5.4), honestly hedged when data is thin. */
function judgeIntent(
  type: PhaseType,
  weightDeltaKg: number | null,
  calorieSkew: 'under' | 'over' | null,
): PhaseIntentVerdict {
  const label = PHASE_TYPE_LABELS[type].toLowerCase();
  const intent = PHASE_INTENT[type];

  if (intent.weight === 'none') {
    return {
      alignment: 'unclear',
      message: `A ${label} block sets no fixed direction — the summary below describes what happened.`,
    };
  }

  if (intent.weight === 'stable') {
    if (weightDeltaKg !== null && Math.abs(weightDeltaKg) > STABLE_DRIFT_KG) {
      return {
        alignment: 'counter',
        message: `Weight moved ${signed(weightDeltaKg)} kg — more drift than a steady ${label} intends.`,
      };
    }
    if (weightDeltaKg !== null && Math.abs(weightDeltaKg) <= WEIGHT_STABLE_KG) {
      return {
        alignment: 'aligned',
        message: `Weight held within ${signed(weightDeltaKg)} kg — steady, as a ${label} intends.`,
      };
    }
    return {
      alignment: 'unclear',
      message: `Not enough weight data to confirm this ${label} stayed steady.`,
    };
  }

  const wantUp = intent.weight === 'up';
  const wrongSkew = wantUp ? 'under' : 'over';
  const rightSkew = wantUp ? 'over' : 'under';
  const movedRight =
    weightDeltaKg !== null &&
    (wantUp ? weightDeltaKg > WEIGHT_STABLE_KG : weightDeltaKg < -WEIGHT_STABLE_KG);
  const movedWrong =
    weightDeltaKg !== null &&
    (wantUp ? weightDeltaKg < -WEIGHT_STABLE_KG : weightDeltaKg > WEIGHT_STABLE_KG);

  if (movedWrong || calorieSkew === wrongSkew) {
    const why = movedWrong
      ? `weight moved ${signed(weightDeltaKg!)} kg`
      : `calories skewed ${calorieSkew} target on most logged days`;
    return { alignment: 'counter', message: `This ${label} ran counter to plan — ${why}.` };
  }
  if (movedRight || calorieSkew === rightSkew) {
    const why = movedRight
      ? `weight moved ${signed(weightDeltaKg!)} kg`
      : `calories skewed ${calorieSkew} target`;
    return { alignment: 'aligned', message: `On track for a ${label}: ${why}.` };
  }
  return {
    alignment: 'unclear',
    message: `Not enough logged data yet to judge this ${label} against its intent.`,
  };
}

/**
 * The Phase analytics calculator (ANALYTICS_ENGINE §5.4), pure. A phase is a **lens over
 * a fixed window** [start, end] — never a bias on other views. Everything is computed
 * strictly inside the phase's own dates (end = today for an ongoing phase), so a
 * completed block always reads the same regardless of what today's phase is. Composes the
 * existing calculators windowed to the phase; adds body deltas, in-phase PRs, and an
 * intent verdict; and normalizes per week for cross-phase comparison.
 */
export function computePhaseReport(input: PhaseReportInput): PhaseReport {
  const { phase, workouts, weighIns, snapshots, nutritionDays, today } = input;
  const start = phase.startDate;
  const ongoing = phase.endDate === null;
  const end: IsoDate = phase.endDate ?? today;
  const spanDays = Math.max(1, daysBetweenIso(start, end) + 1);
  const weeks = spanDays / 7;
  const window: Range = { key: 'all', startDate: start, endDate: end, days: spanDays };

  // Compose the workout calculator, windowed to the phase (today = the phase's own end so
  // its date-relative parts stay inside the block). We consume only its window-scoped
  // fields: total/by-group volume and the key-exercise e1RM trends.
  const workout = computeWorkoutAnalytics({
    workouts,
    weighIns,
    weeklyWorkoutTarget: input.weeklyWorkoutTarget,
    defaultBodyweightKg: input.defaultBodyweightKg,
    scheduledWeekdays: [],
    window,
    today: end,
  });

  const countableDates = workouts
    .filter(
      (w) =>
        inRange(w.date, start, end) &&
        w.exercises.some((ex) => ex.sets.some((s) => isWorkingSet(s, ex.loadType))),
    )
    .map((w) => w.date);

  const body = bodyDeltas(snapshots, start, end, spanDays, input.heightCm);
  const bwAtEnd = resolveBodyweightForDate(weighIns, end, input.defaultBodyweightKg);
  const { prs, prCount } = phasePrs(workouts, start, end, bwAtEnd);

  const nutrition = computeNutritionAnalytics({ days: nutritionDays, window, today: end });

  const weightCmp = isOk(body) ? body.value.fields.find((f) => f.field === 'weightKg') : undefined;
  const weightDeltaKg = weightCmp?.deltaAbs ?? null;

  const training: PhaseTrainingSummary = {
    workouts: countableDates.length,
    weeks: round1(weeks),
    workoutsPerWeek: round1(countableDates.length / weeks),
    avgWeeklyConsistencyPct: avgWeeklyConsistency(
      countableDates,
      start,
      end,
      input.weeklyWorkoutTarget,
    ),
    totalVolumeKg: workout.totalVolumeKg,
    volumeByGroup: workout.volumeByGroup,
    keyExercises: workout.keyExercises,
    prCount,
    prs,
  };

  return {
    phase,
    ongoing,
    window,
    spanDays,
    bodyDeltas: body,
    training,
    nutrition: { nutrition, intent: judgeIntent(phase.type, weightDeltaKg, nutrition.calorieSkew) },
    rates: {
      weeks: round1(weeks),
      workoutsPerWeek: round1(countableDates.length / weeks),
      volumePerWeekKg: Math.round(workout.totalVolumeKg / weeks),
      weightDeltaPerWeekKg: weightDeltaKg !== null ? round1(weightDeltaKg / weeks) : null,
    },
  };
}
