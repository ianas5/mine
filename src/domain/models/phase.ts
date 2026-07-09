import type { EpochMs, IsoDate } from '@/core/utils';

/**
 * A user-declared training phase (DATABASE §3.7, ANALYTICS_ENGINE §5.4). A named
 * period of *intent* — never auto-detected. Phases may not overlap and at most one is
 * ongoing (`endDate === null`); both invariants are enforced in `phaseRepository`.
 *
 * A phase is **context, not prediction**: it explains why a stretch of history looked
 * the way it did. It never rewrites analytics, and historical results are always read
 * against the phase that was active at that time — editing today's phase never
 * reinterprets a past block.
 */
export type PhaseType = 'cutting' | 'recomp' | 'lean_bulk' | 'maintenance' | 'custom';

export const PHASE_TYPES: readonly PhaseType[] = [
  'cutting',
  'recomp',
  'lean_bulk',
  'maintenance',
  'custom',
];

export const PHASE_TYPE_LABELS: Record<PhaseType, string> = {
  cutting: 'Cut',
  recomp: 'Recomp',
  lean_bulk: 'Lean Bulk',
  maintenance: 'Maintenance',
  custom: 'Custom',
};

/** Which way body weight is *meant* to move for a phase type — the yardstick a Phase
 * Report judges the block against (a cut with weight up, or kcal over, is flagged). */
export type WeightIntent = 'down' | 'up' | 'stable' | 'none';

export interface PhaseIntent {
  readonly weight: WeightIntent;
  /** A short plain-English statement of the block's goal (for the report header). */
  readonly summary: string;
}

export const PHASE_INTENT: Record<PhaseType, PhaseIntent> = {
  cutting: { weight: 'down', summary: 'lose fat while holding strength' },
  recomp: { weight: 'stable', summary: 'add muscle and lose fat at a steady weight' },
  lean_bulk: { weight: 'up', summary: 'gain muscle with minimal fat' },
  maintenance: { weight: 'stable', summary: 'hold weight and maintain' },
  custom: { weight: 'none', summary: 'a self-defined block' },
};

/** A declared training phase (DATABASE §3.7). `endDate === null` ⇒ ongoing. */
export interface Phase {
  readonly id: string;
  readonly name: string;
  readonly type: PhaseType;
  readonly startDate: IsoDate;
  /** Inclusive last day of the block, or `null` while the phase is still ongoing. */
  readonly endDate: IsoDate | null;
  readonly notes: string | null;
  readonly createdAt: EpochMs;
  readonly updatedAt: EpochMs;
}
