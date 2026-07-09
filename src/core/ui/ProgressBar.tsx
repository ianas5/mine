import type { ReactNode } from 'react';
import { View } from 'react-native';

import { useTheme } from '@/core/theme';

import type { StatTone } from './StatTile';

interface ProgressBarProps {
  /** 0–1; values outside are clamped (over-target reads as full, never overflowing). */
  readonly value: number;
  readonly tone?: StatTone;
  readonly height?: number;
  readonly accessibilityLabel?: string;
}

function toneColor(theme: ReturnType<typeof useTheme>, tone: StatTone): string {
  switch (tone) {
    case 'positive':
      return theme.color.positive;
    case 'negative':
      return theme.color.danger;
    case 'attention':
      return theme.color.attention;
    case 'neutral':
      return theme.color.accent;
  }
}

/** A macro/consistency progress track (DESIGN_SYSTEM §6): `chartMuted` track, tone fill. */
export function ProgressBar(props: ProgressBarProps): ReactNode {
  const theme = useTheme();
  const pct = Math.max(0, Math.min(1, props.value)) * 100;
  const height = props.height ?? 6;
  return (
    <View
      accessibilityRole="progressbar"
      accessibilityLabel={props.accessibilityLabel}
      style={{
        height,
        borderRadius: theme.radius.full,
        backgroundColor: theme.color.chartMuted,
        overflow: 'hidden',
      }}
    >
      <View
        style={{
          width: `${pct}%`,
          height: '100%',
          borderRadius: theme.radius.full,
          backgroundColor: toneColor(theme, props.tone ?? 'neutral'),
        }}
      />
    </View>
  );
}
