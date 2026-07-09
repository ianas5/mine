import { useEffect, useState, type ReactNode } from 'react';
import { Animated, type ViewStyle } from 'react-native';

import { useReducedMotion } from '@/core/theme';

interface SettleInProps {
  readonly children: ReactNode;
  /** Starting scale for the settle (default 0.85). */
  readonly from?: number;
  readonly style?: ViewStyle;
}

/**
 * A one-shot scale-settle + fade for a freshly-appearing element (delight registry — the
 * PR badge materializing, a workout/phase summary landing). Subtle and ≤ 800 ms; it never
 * gates the underlying write (the element is already there, this only softens its arrival).
 * Under Reduce Motion it appears instantly at full scale/opacity — feedback without motion
 * (UI_UX §9, P15).
 */
export function SettleIn(props: SettleInProps): ReactNode {
  const reduced = useReducedMotion();
  const [scale] = useState(() => new Animated.Value(props.from ?? 0.85));
  const [opacity] = useState(() => new Animated.Value(0));

  useEffect(() => {
    if (reduced) {
      scale.setValue(1);
      opacity.setValue(1);
      return;
    }
    Animated.parallel([
      Animated.spring(scale, { toValue: 1, useNativeDriver: true, friction: 6, tension: 160 }),
      Animated.timing(opacity, { toValue: 1, duration: 160, useNativeDriver: true }),
    ]).start();
  }, [reduced, scale, opacity]);

  return (
    <Animated.View style={{ ...props.style, opacity, transform: [{ scale }] }}>
      {props.children}
    </Animated.View>
  );
}
