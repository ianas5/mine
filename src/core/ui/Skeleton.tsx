import { useEffect, useState, type ReactNode } from 'react';
import { Animated, type DimensionValue } from 'react-native';

import { useReducedMotion, useTheme } from '@/core/theme';

interface SkeletonProps {
  readonly width?: DimensionValue;
  readonly height?: number;
  readonly radius?: number;
}

/** Pulsing first-paint placeholder — never a spinner (P15). Static under Reduce Motion. */
export function Skeleton(props: SkeletonProps): ReactNode {
  const theme = useTheme();
  const [opacity] = useState(() => new Animated.Value(0.5));
  const reduceMotion = useReducedMotion();

  useEffect(() => {
    if (reduceMotion) {
      return;
    }
    const loop = Animated.loop(
      Animated.sequence([
        Animated.timing(opacity, {
          toValue: 0.9,
          duration: theme.motion.slow * 2,
          useNativeDriver: true,
        }),
        Animated.timing(opacity, {
          toValue: 0.5,
          duration: theme.motion.slow * 2,
          useNativeDriver: true,
        }),
      ]),
    );
    loop.start();
    return () => loop.stop();
  }, [opacity, reduceMotion, theme.motion.slow]);

  return (
    <Animated.View
      accessibilityLabel="Loading"
      style={{
        width: props.width ?? '100%',
        height: props.height ?? theme.space.lg,
        borderRadius: props.radius ?? theme.radius.sm,
        backgroundColor: theme.color.border,
        opacity: reduceMotion ? 0.7 : opacity,
      }}
    />
  );
}
