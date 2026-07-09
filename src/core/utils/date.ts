import type { EpochMs, IsoDate } from './scalars';

/** Today's device-local calendar date as `YYYY-MM-DD` (FITNESS_DOMAIN §2.3). */
export function todayIso(now: Date = new Date()): IsoDate {
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

const WEEKDAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'] as const;

/** ISO weekday for a date: 0 = Monday … 6 = Sunday (DATABASE §3.3 templates.weekday). */
export function isoWeekday(iso: IsoDate): number {
  const jsDay = new Date(`${iso}T00:00:00`).getDay(); // 0 = Sunday … 6 = Saturday
  return (jsDay + 6) % 7;
}

/** Full label for an ISO weekday (0 = Monday … 6 = Sunday). */
export function weekdayLabel(isoWeekdayIndex: number): string {
  const labels = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  return labels[isoWeekdayIndex] ?? '';
}

/** Humanized recency: `Today`, `Yesterday`, or `Mon · 4d ago` (FITNESS_DOMAIN §2.3 dates). */
export function formatRelativeDate(iso: IsoDate, now: Date = new Date()): string {
  const d = new Date(`${iso}T00:00:00`);
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const days = Math.round((today.getTime() - d.getTime()) / 86_400_000);
  if (days <= 0) return 'Today';
  if (days === 1) return 'Yesterday';
  return `${WEEKDAYS[d.getDay()] ?? ''} · ${days}d ago`;
}

/** Formats an elapsed duration as `M:SS` (or `H:MM:SS` past an hour). */
export function formatElapsed(ms: EpochMs): string {
  const total = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(total / 3600);
  const m = Math.floor((total % 3600) / 60);
  const s = total % 60;
  const ss = String(s).padStart(2, '0');
  if (h > 0) {
    return `${h}:${String(m).padStart(2, '0')}:${ss}`;
  }
  return `${m}:${ss}`;
}
