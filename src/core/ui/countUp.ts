import { useEffect, useRef, useState } from 'react';

import { useReducedMotion } from '@/core/theme';

/**
 * Eased position of a count-up at progress `t ∈ [0,1]` (easeOutCubic — fast then
 * settling, so a summary number lands rather than races). Pure and clamped; the
 * animation hook is the only impure caller. Extracted so the curve is unit-testable.
 */
export function countUpValue(from: number, to: number, t: number): number {
  const clamped = t <= 0 ? 0 : t >= 1 ? 1 : t;
  const eased = 1 - Math.pow(1 - clamped, 3);
  return from + (to - from) * eased;
}

interface CountUpOptions {
  readonly durationMs?: number;
  /** Animate only while true (e.g. when a summary sheet is visible); default true. */
  readonly enabled?: boolean;
}

/**
 * Counts a number up from 0 to `target` once (delight registry — workout summary, ≤ 800 ms,
 * never blocking). Under Reduce Motion, or when disabled, it returns `target` immediately —
 * the number is always correct, only the approach animates (P15). Drives a 0→1 progress from
 * `requestAnimationFrame` (state is only ever set from the async frame callback, never
 * synchronously in the effect) and derives the displayed value purely, so it never blocks the
 * write path or triggers cascading renders.
 */
export function useCountUp(target: number, options: CountUpOptions = {}): number {
  const reduced = useReducedMotion();
  const durationMs = options.durationMs ?? 450;
  const enabled = options.enabled ?? true;
  const animate = enabled && !reduced && durationMs > 0 && target > 0;
  const [progress, setProgress] = useState(0);
  const rafRef = useRef<number | null>(null);

  useEffect(() => {
    if (!animate) return;
    let start: number | null = null;
    const step = (now: number): void => {
      if (start === null) start = now;
      const t = (now - start) / durationMs;
      setProgress(t >= 1 ? 1 : t);
      if (t < 1) rafRef.current = requestAnimationFrame(step);
    };
    rafRef.current = requestAnimationFrame(step);
    return () => {
      if (rafRef.current !== null) cancelAnimationFrame(rafRef.current);
    };
  }, [animate, durationMs]);

  return animate ? countUpValue(0, target, progress) : target;
}
