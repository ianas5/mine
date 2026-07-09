import { useEffect, useState } from 'react';
import { AccessibilityInfo } from 'react-native';

/**
 * Whether the OS "Reduce Motion" accessibility setting is on (UI_UX §9). Every delight
 * animation reads this and swaps to an instant/cross-fade variant — motion is never the
 * sole carrier of feedback (P15). Subscribes to live changes so a mid-session toggle takes
 * effect without a relaunch.
 */
export function useReducedMotion(): boolean {
  const [reduced, setReduced] = useState(false);
  useEffect(() => {
    let live = true;
    void AccessibilityInfo.isReduceMotionEnabled().then((v) => {
      if (live) setReduced(v);
    });
    const sub = AccessibilityInfo.addEventListener('reduceMotionChanged', setReduced);
    return () => {
      live = false;
      sub.remove();
    };
  }, []);
  return reduced;
}
