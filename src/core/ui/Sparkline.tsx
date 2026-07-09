import type { ReactNode } from 'react';
import { View } from 'react-native';
import Svg, { Polyline } from 'react-native-svg';

import { useTheme } from '@/core/theme';

interface SparklineProps {
  /** Pre-bucketed series values from the engine (components never resample). */
  readonly values: readonly number[];
  readonly width?: number;
  readonly height?: number;
  readonly color?: string;
  readonly accessibilityLabel?: string;
}

/**
 * Inline trend glyph (DESIGN_SYSTEM §6, 60×20, no axes) for report rows. Pure
 * react-native-svg — a flat, axis-less polyline scaled to fit. Fewer than two points
 * render nothing (no fabricated line).
 */
export function Sparkline(props: SparklineProps): ReactNode {
  const theme = useTheme();
  const width = props.width ?? 60;
  const height = props.height ?? 20;
  const color = props.color ?? theme.color.accent;

  if (props.values.length < 2) {
    return <View style={{ width, height }} accessibilityLabel={props.accessibilityLabel} />;
  }

  const min = Math.min(...props.values);
  const max = Math.max(...props.values);
  const span = max - min || 1; // flat series → a centered horizontal line
  const pad = 2;
  const stepX = (width - pad * 2) / (props.values.length - 1);

  const points = props.values
    .map((v, i) => {
      const x = pad + i * stepX;
      const y = pad + (1 - (v - min) / span) * (height - pad * 2);
      return `${x},${y}`;
    })
    .join(' ');

  return (
    <View accessibilityLabel={props.accessibilityLabel}>
      <Svg width={width} height={height}>
        <Polyline
          points={points}
          fill="none"
          stroke={color}
          strokeWidth={1.5}
          strokeLinejoin="round"
          strokeLinecap="round"
        />
      </Svg>
    </View>
  );
}
