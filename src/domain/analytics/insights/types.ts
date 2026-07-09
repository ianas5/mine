import type { IsoDate } from '@/core/utils';
import type { MuscleGroup } from '@/domain/fitness';

import type { BodyAnalytics } from '../bodyAnalytics';
import type { MetricResult, Range } from '../metricResult';
import type { NutritionAnalytics, DailyNutrition } from '../nutritionAnalytics';
import type { RecompSignal } from '../recomp';
import type { WorkoutAnalytics } from '../workoutAnalytics';

export type InsightCategory =
  'training' | 'nutrition' | 'body' | 'consistency' | 'milestone' | 'housekeeping';

export type InsightTone = 'positive' | 'attention' | 'neutral';

/** Where an insight's evidence tap-through lands (UI_UX §8 — one tap from its proof). */
export type InsightEvidence =
  | { readonly kind: 'analytics-body' }
  | { readonly kind: 'analytics-training' }
  | { readonly kind: 'analytics-nutrition' }
  | { readonly kind: 'muscle-report' }
  | { readonly kind: 'measurements' }
  | { readonly kind: 'photos' }
  | { readonly kind: 'exercise'; readonly exerciseId: string };

/** What a rule emits before the engine scores/dedups/cools it (§6.1). */
export interface InsightSeed {
  readonly ruleId: string;
  readonly instanceKey: string;
  readonly category: InsightCategory;
  readonly tone: InsightTone;
  readonly base: number;
  /** Days before the same instanceKey may reappear; `'once'` = once per milestone. */
  readonly cooldownDays: number | 'once';
  readonly title: string;
  readonly body: string;
  readonly data: Record<string, unknown>;
  readonly window: Range;
  /** 0–1: how far past threshold (drives the magnitude bonus). */
  readonly magnitude: number;
  /** Salient classification for flip-breaks-cooldown (e.g. 'improving'); '' if n/a. */
  readonly classification: string;
  readonly evidence: InsightEvidence;
}

/** The scored, selectable insight (§6.1). */
export interface Insight extends InsightSeed {
  readonly priority: number;
}

/** Pre-computed calculator outputs + raw signals the rules read (ANALYTICS §3.4). */
export interface InsightContext {
  readonly today: IsoDate;
  readonly window: Range;
  readonly body: BodyAnalytics;
  readonly recomp: MetricResult<RecompSignal>;
  readonly nutrition: NutritionAnalytics;
  readonly nutritionDays: readonly DailyNutrition[];
  readonly workout: WorkoutAnalytics;
  /** Countable sessions in the last 30 days (gates balance/neglect rules). */
  readonly sessions30d: number;
  /** Completed-week consistency % for the last two fully-elapsed weeks (rules 12/13). */
  readonly completedWeekConsistency: {
    readonly current: number | null;
    readonly previous: number | null;
  };
  /** PR events set in the last 7 days (rule 15). */
  readonly recentPrs: readonly {
    readonly exerciseId: string;
    readonly name: string;
    readonly kinds: readonly string[];
  }[];
  /** Muscle groups with 0 working sets over 30d (rule 19); pre-filtered to canonical. */
  readonly neglectedGroups: readonly MuscleGroup[];
  readonly lastSnapshotDate: IsoDate | null;
  readonly lastPhotoDate: IsoDate | null;
}

/** A rule: a stable id + a pure evaluator. */
export interface InsightRule {
  readonly id: string;
  readonly evaluate: (ctx: InsightContext) => InsightSeed | readonly InsightSeed[] | null;
}

/** Cooldown state (MMKV-persisted): last-fired date + classification per instanceKey. */
export interface CooldownEntry {
  readonly lastFired: IsoDate;
  readonly classification: string;
}
export type CooldownMap = Readonly<Record<string, CooldownEntry>>;
