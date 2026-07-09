import type { IsoDate } from '@/core/utils';

/**
 * The engine result contract (ANALYTICS §3.1). Every metric returns this discriminated
 * union — it never throws for data reasons and never leaks `NaN`/`null` to the UI. A
 * metric is either `ok` (a value plus the window it was computed over) or
 * `insufficient-data` with a concrete, human-readable `needed` sentence.
 */

export type RangeKey = '7d' | '30d' | '90d' | '180d' | '365d' | 'all';

export interface Range {
  readonly key: RangeKey;
  /** Inclusive window start; `null` for all-time with no data. */
  readonly startDate: IsoDate | null;
  readonly endDate: IsoDate;
  /** Fixed window length in days; `null` for all-time (open-ended start). */
  readonly days: number | null;
}

export type InsufficientReason = 'no-data' | 'too-few-points' | 'span-too-short' | 'no-target-set';

export interface ComputedFrom {
  readonly points: number;
  readonly spanDays: number;
}

export type MetricResult<T> =
  | {
      readonly status: 'ok';
      readonly value: T;
      readonly window: Range;
      readonly computedFrom: ComputedFrom;
    }
  | {
      readonly status: 'insufficient-data';
      readonly reason: InsufficientReason;
      readonly needed: string;
    };

export function ok<T>(value: T, window: Range, computedFrom: ComputedFrom): MetricResult<T> {
  return { status: 'ok', value, window, computedFrom };
}

export function insufficient<T = never>(
  reason: InsufficientReason,
  needed: string,
): MetricResult<T> {
  return { status: 'insufficient-data', reason, needed };
}

export function isOk<T>(
  result: MetricResult<T>,
): result is { status: 'ok'; value: T; window: Range; computedFrom: ComputedFrom } {
  return result.status === 'ok';
}
