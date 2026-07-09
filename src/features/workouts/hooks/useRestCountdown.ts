import { useEffect, useRef, useState } from 'react';

import { triggerHaptic } from '@/core/theme';

import { useRestTimerStore } from '../stores/useRestTimerStore';

export interface RestCountdown {
  readonly running: boolean;
  readonly remainingMs: number;
  readonly durationSec: number;
}

/**
 * Live rest countdown derived from the wall-clock `endsAt` (honest across
 * background/foreground — no drift). Fires a single gentle haptic and clears the
 * timer the moment it reaches zero (UI_UX §4.2). Ticking is derived, never a
 * synchronous setState in an effect.
 */
export function useRestCountdown(): RestCountdown {
  const running = useRestTimerStore((s) => s.running);
  const endsAt = useRestTimerStore((s) => s.endsAt);
  const durationSec = useRestTimerStore((s) => s.durationSec);
  const skip = useRestTimerStore((s) => s.actions.skip);

  const [now, setNow] = useState(() => Date.now());

  useEffect(() => {
    if (!running || endsAt === null) return;
    const id = setInterval(() => setNow(Date.now()), 500);
    return () => clearInterval(id);
  }, [running, endsAt]);

  // Clamp to the duration so a stale `now` (carried from a previous rest) can
  // never briefly show more than the full window before the first tick lands.
  const remainingMs =
    running && endsAt !== null ? Math.min(durationSec * 1000, Math.max(0, endsAt - now)) : 0;

  // Completion fires exactly once per rest window (keyed on endsAt).
  const firedFor = useRef<number | null>(null);
  useEffect(() => {
    if (running && endsAt !== null && remainingMs <= 0 && firedFor.current !== endsAt) {
      firedFor.current = endsAt;
      triggerHaptic('success');
      skip();
    }
  }, [running, endsAt, remainingMs, skip]);

  return { running, remainingMs, durationSec };
}

/** mm:ss for a rest countdown (ceil so it reads 1:30 → 1:29 → … → 0:01 → done). */
export function formatCountdown(remainingMs: number): string {
  const totalSec = Math.ceil(remainingMs / 1000);
  const minutes = Math.floor(totalSec / 60);
  const seconds = totalSec % 60;
  return `${minutes}:${seconds.toString().padStart(2, '0')}`;
}
