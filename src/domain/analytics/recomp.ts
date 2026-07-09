import { addDaysIso, daysBetweenIso, type IsoDate } from '@/core/utils';
import type { BodySnapshot } from '@/domain/body';

import { insufficient, ok, type MetricResult, type Range } from './metricResult';

/** §6.5 windows + thresholds (verbatim). */
const WINDOW_DAYS = 56; // 8 weeks
const MIN_SPAN_DAYS = 28; // minimum 4 weeks
const WEIGHT_STABLE_KG = 1.0;
const WEIGHT_STABLE_PCT = 1.5;
const WAIST_DOWN_CM = 1.0;
const BODYFAT_DOWN_PCT = 0.5;
const MUSCLE_UP_KG = 0.3;

export interface RecompSignal {
  readonly fired: boolean;
  readonly weightDeltaKg: number;
  readonly waistDeltaCm: number | null;
  readonly bodyFatDeltaPct: number | null;
  readonly muscleDeltaKg: number | null;
  readonly spanDays: number;
  /** Human-readable markers that fired (for the insight body). */
  readonly markers: readonly string[];
}

const delta = (a: number | null, b: number | null): number | null =>
  a !== null && b !== null ? b - a : null;

/**
 * Body recomposition signal (FITNESS_DOMAIN §6.5), pure. Over the trailing 8-week window
 * (min 4 weeks, using the earliest and latest snapshots in it), fires when weight is stable
 * (|Δ| ≤ 1 kg **or** |Δ%| ≤ 1.5%) **and** at least one fat-down/muscle-up marker holds. A
 * meaningful weight drop alongside a waist/fat drop is *fat loss*, not recomp — the
 * stability condition guards against that. Returns `insufficient-data` below the minimums.
 */
export function computeRecompSignal(
  snapshots: readonly BodySnapshot[],
  today: IsoDate,
): MetricResult<RecompSignal> {
  const start = addDaysIso(today, -(WINDOW_DAYS - 1));
  const inWindow = snapshots
    .filter((s) => s.date >= start && s.date <= today && s.weightKg !== null)
    .sort((a, b) => (a.date < b.date ? -1 : 1));
  const window: Range = { key: 'all', startDate: start, endDate: today, days: WINDOW_DAYS };

  if (inWindow.length < 2) {
    return insufficient(
      'too-few-points',
      'Log two weigh-ins about 8 weeks apart to check for recomposition',
    );
  }
  const first = inWindow[0]!;
  const last = inWindow[inWindow.length - 1]!;
  const spanDays = daysBetweenIso(first.date, last.date);
  if (spanDays < MIN_SPAN_DAYS) {
    return insufficient(
      'span-too-short',
      'A recomposition read needs at least 4 weeks between weigh-ins',
    );
  }

  const weightDeltaKg = last.weightKg! - first.weightKg!;
  const weightPct = first.weightKg! !== 0 ? (weightDeltaKg / first.weightKg!) * 100 : 0;
  const weightStable =
    Math.abs(weightDeltaKg) <= WEIGHT_STABLE_KG || Math.abs(weightPct) <= WEIGHT_STABLE_PCT;

  const waistDeltaCm = delta(first.waistCm, last.waistCm);
  const bodyFatDeltaPct = delta(first.bodyFatPct, last.bodyFatPct);
  const muscleDeltaKg = delta(first.muscleMassKg, last.muscleMassKg);

  const markers: string[] = [];
  if (waistDeltaCm !== null && waistDeltaCm <= -WAIST_DOWN_CM) {
    markers.push(`waist ${waistDeltaCm.toFixed(1)} cm`);
  }
  if (bodyFatDeltaPct !== null && bodyFatDeltaPct <= -BODYFAT_DOWN_PCT) {
    markers.push(`body fat ${bodyFatDeltaPct.toFixed(1)}%`);
  }
  if (muscleDeltaKg !== null && muscleDeltaKg >= MUSCLE_UP_KG) {
    markers.push(`muscle +${muscleDeltaKg.toFixed(1)} kg`);
  }

  return ok(
    {
      fired: weightStable && markers.length > 0,
      weightDeltaKg,
      waistDeltaCm,
      bodyFatDeltaPct,
      muscleDeltaKg,
      spanDays,
      markers,
    },
    window,
    { points: inWindow.length, spanDays },
  );
}
