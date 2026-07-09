import type { ReactNode } from 'react';
import { Text, View } from 'react-native';
import Svg, { Circle } from 'react-native-svg';

import { useTheme } from '@/core/theme';
import type { StatTone } from '@/core/ui';

interface RingProps {
  readonly label: string;
  /** Center value, pre-formatted (e.g. "820"). */
  readonly value: string;
  readonly unit: string;
  /** Consumed ÷ target, 0–1 (clamped); drives the arc fill. */
  readonly fraction: number;
  readonly tone?: StatTone;
  readonly size?: number;
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

/** A single macro ring (DESIGN_SYSTEM §6.1): the arc shows progress-to-target, the
 * center shows what's left. react-native-svg (device-safe). */
export function Ring(props: RingProps): ReactNode {
  const theme = useTheme();
  const size = props.size ?? 96;
  const stroke = 8;
  const r = (size - stroke) / 2;
  const c = 2 * Math.PI * r;
  const filled = Math.max(0, Math.min(1, props.fraction));

  return (
    <View style={{ alignItems: 'center', gap: theme.space.xs }}>
      <View style={{ width: size, height: size, alignItems: 'center', justifyContent: 'center' }}>
        <Svg width={size} height={size} style={{ position: 'absolute' }}>
          <Circle
            cx={size / 2}
            cy={size / 2}
            r={r}
            stroke={theme.color.chartMuted}
            strokeWidth={stroke}
            fill="none"
          />
          <Circle
            cx={size / 2}
            cy={size / 2}
            r={r}
            stroke={toneColor(theme, props.tone ?? 'neutral')}
            strokeWidth={stroke}
            fill="none"
            strokeLinecap="round"
            strokeDasharray={c}
            strokeDashoffset={c * (1 - filled)}
            transform={`rotate(-90 ${size / 2} ${size / 2})`}
          />
        </Svg>
        <Text
          style={{
            ...theme.type.heading,
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {props.value}
        </Text>
        <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>{props.unit}</Text>
      </View>
      <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>{props.label}</Text>
    </View>
  );
}
