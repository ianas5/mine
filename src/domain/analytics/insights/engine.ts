import { addDaysIso } from '@/core/utils';

import { RULES } from './rules';
import type { CooldownMap, Insight, InsightContext, InsightSeed } from './types';

const DASHBOARD_MAX = 3;
const MAX_PER_CATEGORY = 2;
const MAX_HOUSEKEEPING = 1;
const FRESHNESS_BONUS = 10; // v1: every insight is evaluated fresh, so this is constant

/** Score = base + magnitude bonus (0–10) + freshness bonus (§6.3). */
function score(seed: InsightSeed): number {
  return seed.base + Math.round(seed.magnitude * 10) + FRESHNESS_BONUS;
}

/**
 * A fired instance is suppressed inside its cooldown **unless its classification flips**
 * (§6.3 — direction changes are never suppressed). `'once'` cooldown never re-fires.
 */
function withinCooldown(seed: InsightSeed, cooldowns: CooldownMap, today: string): boolean {
  const entry = cooldowns[seed.instanceKey];
  if (!entry) return false;
  if (seed.classification !== '' && seed.classification !== entry.classification) return false; // flip
  if (seed.cooldownDays === 'once') return true;
  return entry.lastFired > addDaysIso(today, -seed.cooldownDays);
}

/**
 * Evaluates all §6.2 rules into the live, de-duplicated, cooled-down, scored insight list
 * (highest first). Pure: cooldown state is passed in (MMKV-persisted by the caller). Mutually
 * exclusive rules can't co-fire because their triggers negate each other (§6.3 conflict guard).
 */
export function evaluateInsights(ctx: InsightContext, cooldowns: CooldownMap = {}): Insight[] {
  const seen = new Set<string>();
  const insights: Insight[] = [];

  for (const rule of RULES) {
    const result = rule.evaluate(ctx);
    if (result === null) continue;
    const seeds = Array.isArray(result) ? result : [result as InsightSeed];
    for (const seed of seeds) {
      if (seen.has(seed.instanceKey)) continue; // dedup: one live insight per instanceKey
      seen.add(seed.instanceKey);
      if (withinCooldown(seed, cooldowns, ctx.today)) continue;
      insights.push({ ...seed, priority: score(seed) });
    }
  }

  return insights.sort((a, b) => b.priority - a.priority);
}

/**
 * Dashboard selection (§6.3): at most 3, highest score first, ≤ 2 of any category and
 * ≤ 1 housekeeping. The Analytics tab shows the full list; this is only the dashboard cut.
 */
export function selectDashboardInsights(insights: readonly Insight[]): Insight[] {
  const chosen: Insight[] = [];
  const perCategory = new Map<string, number>();
  let housekeeping = 0;

  for (const insight of insights) {
    if (chosen.length >= DASHBOARD_MAX) break;
    if ((perCategory.get(insight.category) ?? 0) >= MAX_PER_CATEGORY) continue;
    if (insight.category === 'housekeeping' && housekeeping >= MAX_HOUSEKEEPING) continue;
    chosen.push(insight);
    perCategory.set(insight.category, (perCategory.get(insight.category) ?? 0) + 1);
    if (insight.category === 'housekeeping') housekeeping += 1;
  }
  return chosen;
}

/**
 * The cooldown-map update after a set of insights is shown/fired: stamp each shown
 * instanceKey with today + its classification (for later flip detection).
 */
export function stampCooldowns(
  shown: readonly Insight[],
  cooldowns: CooldownMap,
  today: string,
): CooldownMap {
  const next: Record<string, { lastFired: string; classification: string }> = { ...cooldowns };
  for (const insight of shown) {
    next[insight.instanceKey] = { lastFired: today, classification: insight.classification };
  }
  return next;
}
