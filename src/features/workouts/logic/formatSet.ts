import { formatKg } from '@/core/utils';
import type { LoadType } from '@/domain/fitness';

/** Compact set display per load type: `80 × 8`, `8`, or `45s`. Pure. */
export function formatSet(weightKg: number, reps: number, loadType: LoadType): string {
  if (loadType === 'timed') return `${reps}s`;
  if (loadType === 'bodyweight') return `${reps}`;
  return `${formatKg(weightKg)} × ${reps}`;
}
