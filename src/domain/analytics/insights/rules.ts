import type { BodyField } from '@/domain/body';

import { isOk } from '../metricResult';
import { calorieSkew, proteinMissStreak, trailingSignals } from '../nutritionAnalytics';
import type { InsightContext, InsightEvidence, InsightRule, InsightSeed } from './types';

/**
 * The v1 insight rule catalog (ANALYTICS §6.2), verbatim to the table. Each rule is a pure
 * evaluator that connects the calculators' numbers into a worded "so what?" — actionable,
 * specific, evidence-based, time-aware, calm, hedged where inferential (§6.1 wording rules).
 * Adding a rule means amending §6.2 first, then this list.
 */

const clamp01 = (n: number): number => Math.max(0, Math.min(1, n));

interface SeedArgs {
  ruleId: string;
  instanceKey?: string;
  category: InsightSeed['category'];
  tone: InsightSeed['tone'];
  base: number;
  cooldownDays: number | 'once';
  title: string;
  body: string;
  evidence: InsightEvidence;
  magnitude?: number;
  classification?: string;
  data?: Record<string, unknown>;
  window: InsightSeed['window'];
}

const seed = (a: SeedArgs): InsightSeed => ({
  ruleId: a.ruleId,
  instanceKey: a.instanceKey ?? a.ruleId,
  category: a.category,
  tone: a.tone,
  base: a.base,
  cooldownDays: a.cooldownDays,
  title: a.title,
  body: a.body,
  data: a.data ?? {},
  window: a.window,
  magnitude: clamp01(a.magnitude ?? 0.5),
  classification: a.classification ?? '',
  evidence: a.evidence,
});

const SITE_LABEL: Partial<Record<BodyField, string>> = {
  waistCm: 'waist',
  chestCm: 'chest',
  leftArmCm: 'left arm',
  rightArmCm: 'right arm',
  leftThighCm: 'left thigh',
  rightThighCm: 'right thigh',
};
const MUSCULAR_SITES: readonly BodyField[] = [
  'chestCm',
  'leftArmCm',
  'rightArmCm',
  'leftThighCm',
  'rightThighCm',
];

const kg1 = (n: number): string => (Math.round(n * 10) / 10).toString();

export const RULES: readonly InsightRule[] = [
  // 1 — recomposition (the prime signal)
  {
    id: 'recomp-signal',
    evaluate: (ctx) => {
      if (!isOk(ctx.recomp) || !ctx.recomp.value.fired) return null;
      const r = ctx.recomp.value;
      return seed({
        ruleId: 'recomp-signal',
        category: 'body',
        tone: 'positive',
        base: 95,
        cooldownDays: 14,
        title: 'Possible body recomposition',
        body: `Your body weight held steady while ${r.markers.join(' and ')} moved over ${Math.round(r.spanDays / 7)} weeks. This may indicate you're losing fat while keeping muscle.`,
        evidence: { kind: 'analytics-body' },
        magnitude: 1,
        classification: 'recomp',
        data: { markers: r.markers, spanDays: r.spanDays },
        window: ctx.recomp.window,
      });
    },
  },

  // 2 — weight trending toward goal
  {
    id: 'weight-trend-to-target',
    evaluate: (ctx) => {
      const t = ctx.body.weightTrend;
      const d = ctx.body.distanceToTarget;
      if (!isOk(t) || t.value.classification !== 'improving') return null;
      if (!isOk(d) || d.value.atGoal) return null;
      return seed({
        ruleId: 'weight-trend-to-target',
        category: 'body',
        tone: 'positive',
        base: 70,
        cooldownDays: 7,
        title: 'On track to your goal weight',
        body: `Your trend weight is moving toward your goal at ${kg1(Math.abs(t.value.slopePerWeek))} kg/week — about ${kg1(Math.abs(d.value.toGoKg))} kg to go${d.value.etaWeeks ? `, ~${d.value.etaWeeks} weeks at this pace` : ''}.`,
        evidence: { kind: 'analytics-body' },
        classification: 'improving',
        magnitude: clamp01(Math.abs(t.value.slopePerWeek) / 0.7),
        window: t.window,
      });
    },
  },

  // 3 — weight moving away from goal
  {
    id: 'weight-trend-away',
    evaluate: (ctx) => {
      const t = ctx.body.weightTrend;
      if (!isOk(t) || t.value.classification !== 'declining') return null;
      return seed({
        ruleId: 'weight-trend-away',
        category: 'body',
        tone: 'attention',
        base: 80,
        cooldownDays: 7,
        title: 'Weight is drifting from your goal',
        body: `Your trend weight has moved away from your goal by ${kg1(Math.abs(t.value.deltaOverWindow))} kg over the window. Worth a look at the last few weeks of intake.`,
        evidence: { kind: 'analytics-body' },
        classification: 'declining',
        magnitude: clamp01(Math.abs(t.value.slopePerWeek) / 0.7),
        window: t.window,
      });
    },
  },

  // 4 — weight stalled far from goal with low kcal adherence
  {
    id: 'weight-stalled',
    evaluate: (ctx) => {
      const t = ctx.body.weightTrend;
      const d = ctx.body.distanceToTarget;
      const cal = ctx.nutrition.calorieAdherence;
      if (!isOk(t) || t.value.direction !== 'stable' || t.computedFrom.spanDays < 28) return null;
      if (!isOk(d) || Math.abs(d.value.toGoKg) <= 2) return null;
      if (!isOk(cal) || cal.value.pct >= 60) return null;
      return seed({
        ruleId: 'weight-stalled',
        category: 'body',
        tone: 'attention',
        base: 75,
        cooldownDays: 14,
        title: 'Weight has stalled short of your goal',
        body: `Your weight has been flat for 4+ weeks while still ${kg1(Math.abs(d.value.toGoKg))} kg from goal, and calories landed on target only ${cal.value.pct}% of logged days. Tightening intake consistency may restart progress.`,
        evidence: { kind: 'analytics-nutrition' },
        magnitude: clamp01((60 - cal.value.pct) / 60),
        window: t.window,
      });
    },
  },

  // 5 — waist decreasing
  {
    id: 'waist-decreasing',
    evaluate: (ctx) => {
      const w = ctx.body.siteTrends.get('waistCm')?.trend;
      if (!w || !isOk(w) || w.value.classification !== 'improving') return null;
      return seed({
        ruleId: 'waist-decreasing',
        category: 'body',
        tone: 'positive',
        base: 75,
        cooldownDays: 14,
        title: 'Your waist is trending down',
        body: `Waist measurements are decreasing at about ${kg1(Math.abs(w.value.slopePerWeek))} cm/week — a good sign for body composition.`,
        evidence: { kind: 'analytics-body' },
        classification: 'improving',
        magnitude: clamp01(Math.abs(w.value.slopePerWeek) / 0.5),
        window: w.window,
      });
    },
  },

  // 6 — muscular site growing (per site)
  {
    id: 'site-growing',
    evaluate: (ctx) =>
      MUSCULAR_SITES.flatMap((field) => {
        const s = ctx.body.siteTrends.get(field)?.trend;
        if (!s || !isOk(s) || s.value.classification !== 'improving') return [];
        return [
          seed({
            ruleId: 'site-growing',
            instanceKey: `site-growing:${field}`,
            category: 'body',
            tone: 'positive',
            base: 65,
            cooldownDays: 14,
            title: `Your ${SITE_LABEL[field]} is growing`,
            body: `Your ${SITE_LABEL[field]} has been trending up over the last few months — the training is showing.`,
            evidence: { kind: 'analytics-body' },
            classification: 'improving',
            magnitude: clamp01(Math.abs(s.value.slopePerWeek) / 0.5),
            window: s.window,
          }),
        ];
      }),
  },

  // 7 — muscular site declining (per site)
  {
    id: 'site-declining',
    evaluate: (ctx) =>
      MUSCULAR_SITES.flatMap((field) => {
        const s = ctx.body.siteTrends.get(field)?.trend;
        if (!s || !isOk(s) || s.value.classification !== 'declining') return [];
        return [
          seed({
            ruleId: 'site-declining',
            instanceKey: `site-declining:${field}`,
            category: 'body',
            tone: 'attention',
            base: 70,
            cooldownDays: 14,
            title: `Your ${SITE_LABEL[field]} is shrinking`,
            body: `Your ${SITE_LABEL[field]} has trended down over the last few months. If muscle is the goal, it may be worth checking volume here.`,
            evidence: { kind: 'muscle-report' },
            classification: 'declining',
            magnitude: clamp01(Math.abs(s.value.slopePerWeek) / 0.5),
            window: s.window,
          }),
        ];
      }),
  },

  // 8 — protein miss streak
  {
    id: 'protein-miss-streak',
    evaluate: (ctx) => {
      const streak = protMissStreak(ctx);
      if (streak < 3) return null;
      return seed({
        ruleId: 'protein-miss-streak',
        category: 'nutrition',
        tone: 'attention',
        base: 85,
        cooldownDays: 3,
        title: 'Protein has been short',
        body: `Protein landed below target ${streak} logged days in a row. A protein-forward first meal tomorrow would break the run.`,
        evidence: { kind: 'analytics-nutrition' },
        magnitude: clamp01((streak - 2) / 5),
        data: { streak },
        window: ctx.window,
      });
    },
  },

  // 9 — protein strong week
  {
    id: 'protein-strong-week',
    evaluate: (ctx) => {
      const s = trailing(ctx, 7);
      if (s.loggedCount < 5 || s.proteinHits < 6) return null;
      return seed({
        ruleId: 'protein-strong-week',
        category: 'nutrition',
        tone: 'positive',
        base: 55,
        cooldownDays: 7,
        title: 'Strong protein week',
        body: `You hit your protein target ${s.proteinHits} of the last ${s.loggedCount} logged days. That consistency is what drives muscle retention.`,
        evidence: { kind: 'analytics-nutrition' },
        magnitude: clamp01(s.proteinHits / 7),
        window: ctx.window,
      });
    },
  },

  // 10 — calorie skew (chronic over/under)
  {
    id: 'kcal-skew',
    evaluate: (ctx) => {
      const skew = skew14(ctx);
      if (skew === null) return null;
      return seed({
        ruleId: 'kcal-skew',
        category: 'nutrition',
        tone: 'attention',
        base: 70,
        cooldownDays: 10,
        title: skew === 'over' ? 'Calories running high' : 'Calories running low',
        body:
          skew === 'over'
            ? 'Most of your recent calorie misses were over target. If fat loss is the goal, this is the lever to pull.'
            : 'Most of your recent calorie misses were under target. Chronic under-eating can stall progress and recovery.',
        evidence: { kind: 'analytics-nutrition' },
        classification: skew,
        magnitude: 0.7,
        window: ctx.window,
      });
    },
  },

  // 11 — logging gap
  {
    id: 'logging-gap',
    evaluate: (ctx) => {
      const s = trailing(ctx, 7);
      if (s.loggedCount >= 4) return null;
      return seed({
        ruleId: 'logging-gap',
        category: 'nutrition',
        tone: 'neutral',
        base: 50,
        cooldownDays: 7,
        title: 'Nutrition logging has slipped',
        body: `You logged ${s.loggedCount} of the last 7 days. A couple more logged days would make your nutrition trends reliable again.`,
        evidence: { kind: 'analytics-nutrition' },
        magnitude: clamp01((4 - s.loggedCount) / 4),
        window: ctx.window,
      });
    },
  },

  // 12 — consistency up
  {
    id: 'consistency-up',
    evaluate: (ctx) => {
      const { current, previous } = ctx.completedWeekConsistency;
      if (current === null || previous === null || current <= previous) return null;
      return seed({
        ruleId: 'consistency-up',
        category: 'consistency',
        tone: 'positive',
        base: 60,
        cooldownDays: 7,
        title: 'Training consistency improved',
        body: `Last completed week you hit ${current}% of planned sessions, up from ${previous}% the week before.`,
        evidence: { kind: 'analytics-training' },
        classification: 'up',
        magnitude: clamp01((current - previous) / 50),
        window: ctx.window,
      });
    },
  },

  // 13 — consistency down
  {
    id: 'consistency-down',
    evaluate: (ctx) => {
      const { current, previous } = ctx.completedWeekConsistency;
      if (current === null || previous === null || previous - current < 25) return null;
      return seed({
        ruleId: 'consistency-down',
        category: 'consistency',
        tone: 'attention',
        base: 75,
        cooldownDays: 7,
        title: 'Training consistency dropped',
        body: `Last completed week you hit ${current}% of planned sessions, down from ${previous}% the week before. One session back on the calendar resets the trend.`,
        evidence: { kind: 'analytics-training' },
        classification: 'down',
        magnitude: clamp01((previous - current) / 75),
        window: ctx.window,
      });
    },
  },

  // 14 — streak milestone
  {
    id: 'streak-milestone',
    evaluate: (ctx) => {
      const streak = ctx.workout.consistency.streak;
      if (![4, 8, 12, 26, 52].includes(streak)) return null;
      return seed({
        ruleId: 'streak-milestone',
        instanceKey: `streak-milestone:${streak}`,
        category: 'milestone',
        tone: 'positive',
        base: 65,
        cooldownDays: 'once',
        title: `${streak}-week training streak`,
        body: `You've met your weekly target ${streak} weeks running. Consistency like this is the whole game.`,
        evidence: { kind: 'analytics-training' },
        magnitude: clamp01(streak / 52),
        window: ctx.window,
      });
    },
  },

  // 15 — new PR (per exercise)
  {
    id: 'new-pr',
    evaluate: (ctx) =>
      ctx.recentPrs.map((pr) =>
        seed({
          ruleId: 'new-pr',
          instanceKey: `new-pr:${pr.exerciseId}:${[...pr.kinds].sort().join(',')}`,
          category: 'milestone',
          tone: 'positive',
          base: 80,
          cooldownDays: 30,
          title: `New PR: ${pr.name}`,
          body: `You set a new ${prLabel(pr.kinds)} on ${pr.name} in the last week. Nice work.`,
          evidence: { kind: 'exercise', exerciseId: pr.exerciseId },
          magnitude: 0.8,
          window: ctx.window,
        }),
      ),
  },

  // 16 — strength trend up (per key exercise)
  {
    id: 'strength-trend-up',
    evaluate: (ctx) =>
      ctx.workout.keyExercises.flatMap((k) => {
        if (!isOk(k.trend) || k.trend.value.classification !== 'improving') return [];
        return [
          seed({
            ruleId: 'strength-trend-up',
            instanceKey: `strength-trend-up:${k.exerciseId}`,
            category: 'training',
            tone: 'positive',
            base: 70,
            cooldownDays: 21,
            title: `${k.name} is getting stronger`,
            body: `Your estimated 1RM on ${k.name} has climbed about ${kg1(k.trend.value.slopePerWeek)} kg/week. Progressive overload is working here.`,
            evidence: { kind: 'exercise', exerciseId: k.exerciseId },
            classification: 'improving',
            magnitude: clamp01(k.trend.value.slopePerWeek / 3),
            window: k.trend.window,
          }),
        ];
      }),
  },

  // 17 — strength trend down (per key exercise)
  {
    id: 'strength-trend-down',
    evaluate: (ctx) =>
      ctx.workout.keyExercises.flatMap((k) => {
        if (!isOk(k.trend) || k.trend.value.classification !== 'declining') return [];
        return [
          seed({
            ruleId: 'strength-trend-down',
            instanceKey: `strength-trend-down:${k.exerciseId}`,
            category: 'training',
            tone: 'attention',
            base: 75,
            cooldownDays: 21,
            title: `${k.name} strength is slipping`,
            body: `Your estimated 1RM on ${k.name} has trended down lately. A deload or a form/recovery check may be worth it.`,
            evidence: { kind: 'exercise', exerciseId: k.exerciseId },
            classification: 'declining',
            magnitude: clamp01(Math.abs(k.trend.value.slopePerWeek) / 3),
            window: k.trend.window,
          }),
        ];
      }),
  },

  // 18 — push:pull imbalance
  {
    id: 'push-pull-imbalance',
    evaluate: (ctx) => {
      if (
        ctx.sessions30d < 8 ||
        !ctx.workout.pushPull.flagged ||
        ctx.workout.pushPull.ratio === null
      ) {
        return null;
      }
      const r = ctx.workout.pushPull.ratio;
      return seed({
        ruleId: 'push-pull-imbalance',
        category: 'training',
        tone: 'attention',
        base: 65,
        cooldownDays: 14,
        title: r > 1.25 ? 'Push volume outweighs pull' : 'Pull volume outweighs push',
        body:
          r > 1.25
            ? `Your push-to-pull volume ratio is ${r.toFixed(1)} over the last 30 days. Consider adding pull work to even it out.`
            : `Your push-to-pull volume ratio is ${r.toFixed(1)} over the last 30 days. Consider adding push work to even it out.`,
        evidence: { kind: 'analytics-training' },
        magnitude: clamp01(Math.abs(Math.log(r))),
        window: ctx.window,
      });
    },
  },

  // 19 — neglected muscle (per group)
  {
    id: 'neglected-muscle',
    evaluate: (ctx) => {
      if (ctx.sessions30d < 8) return null;
      return ctx.neglectedGroups.map((group) =>
        seed({
          ruleId: 'neglected-muscle',
          instanceKey: `neglected-muscle:${group}`,
          category: 'training',
          tone: 'attention',
          base: 60,
          cooldownDays: 14,
          title: `${cap(group)} hasn't been trained`,
          body: `You've trained ${ctx.sessions30d} times in 30 days but logged no working sets for ${group}. A little direct work would round things out.`,
          evidence: { kind: 'muscle-report' },
          magnitude: 0.6,
          data: { group },
          window: ctx.window,
        }),
      );
    },
  },

  // 20 — volume drop
  {
    id: 'volume-drop',
    evaluate: (ctx) => {
      const drop = volumeDrop(ctx);
      if (drop === null) return null;
      return seed({
        ruleId: 'volume-drop',
        category: 'training',
        tone: 'attention',
        base: 65,
        cooldownDays: 14,
        title: 'Training volume dropped off',
        body: `Last week's training volume was ${Math.round(drop.pct)}% of your recent 4-week average. If it wasn't a planned deload, it's worth getting back on track.`,
        evidence: { kind: 'analytics-training' },
        magnitude: clamp01((60 - drop.pct) / 60),
        window: ctx.window,
      });
    },
  },

  // 21 — measurement due
  {
    id: 'measurement-due',
    evaluate: (ctx) => {
      const days = daysSince(ctx.lastSnapshotDate, ctx.today);
      if (days !== null && days < 21) return null;
      return seed({
        ruleId: 'measurement-due',
        category: 'housekeeping',
        tone: 'neutral',
        base: 40,
        cooldownDays: 7,
        title: 'Time for measurements',
        body:
          days === null
            ? 'No measurements logged yet — a first set gives your body trends something to work with.'
            : `It's been ${days} days since your last measurements. A quick update keeps your body trends current.`,
        evidence: { kind: 'measurements' },
        magnitude: 0.4,
        window: ctx.window,
      });
    },
  },

  // 22 — photo due
  {
    id: 'photo-due',
    evaluate: (ctx) => {
      const days = daysSince(ctx.lastPhotoDate, ctx.today);
      if (days !== null && days < 35) return null;
      return seed({
        ruleId: 'photo-due',
        category: 'housekeeping',
        tone: 'neutral',
        base: 35,
        cooldownDays: 7,
        title: 'Time for a progress photo',
        body:
          days === null
            ? 'No progress photos yet — the first one becomes your baseline for visual change.'
            : `It's been ${days} days since your last progress photo. Photos catch changes the scale misses.`,
        evidence: { kind: 'photos' },
        magnitude: 0.35,
        window: ctx.window,
      });
    },
  },

  // 23 — water low week
  {
    id: 'water-low-week',
    evaluate: (ctx) => {
      const s = trailing(ctx, 7);
      if (s.waterTargetDays === 0 || s.waterHits >= 3) return null;
      return seed({
        ruleId: 'water-low-week',
        category: 'nutrition',
        tone: 'neutral',
        base: 45,
        cooldownDays: 7,
        title: 'Water has been low',
        body: `You hit your water goal ${s.waterHits} of the last ${s.waterTargetDays} logged days. Front-loading water earlier in the day tends to help.`,
        evidence: { kind: 'analytics-nutrition' },
        magnitude: 0.45,
        window: ctx.window,
      });
    },
  },
];

// --- small helpers the rules lean on (kept here so the catalog reads top-to-bottom) ---

const cap = (s: string): string => s.charAt(0).toUpperCase() + s.slice(1);

function prLabel(kinds: readonly string[]): string {
  const map: Record<string, string> = {
    weight: 'weight',
    e1rm: 'estimated-1RM',
    setVolume: 'set-volume',
    sessionVolume: 'session-volume',
    repAtLoad: 'reps-at-load',
  };
  return kinds.map((k) => map[k] ?? k).join(' & ') + ' PR';
}

function daysSince(date: string | null, today: string): number | null {
  if (date === null) return null;
  const a = new Date(`${date}T00:00:00`).getTime();
  const b = new Date(`${today}T00:00:00`).getTime();
  return Math.round((b - a) / 86_400_000);
}

const protMissStreak = (ctx: InsightContext): number => proteinMissStreak(ctx.nutritionDays);
const trailing = (ctx: InsightContext, n: number) =>
  trailingSignals(ctx.nutritionDays, ctx.today, n);
const skew14 = (ctx: InsightContext) => calorieSkew(ctx.nutritionDays, ctx.today, 14);

function volumeDrop(ctx: InsightContext): { pct: number } | null {
  const series = ctx.workout.volumeSeries;
  if (series.length < 5) return null;
  // Last entry is the most recent completed/there week; compare to the prior 4-week mean.
  const last = series[series.length - 1]!;
  const prior = series.slice(-5, -1);
  if (prior.length < 4) return null;
  const avg = prior.reduce((s, p) => s + p.value, 0) / prior.length;
  if (avg <= 0) return null;
  const pct = (last.value / avg) * 100;
  return pct < 60 ? { pct } : null;
}
