import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';

import type { StatTone } from './StatTile';

interface ChartFrameProps {
  readonly title: string;
  /**
   * The **required** coach sentence for this chart (P6/P8): the chart is never the sole
   * carrier of a conclusion, so a plot without interpretation cannot be rendered.
   */
  readonly interpretation: string;
  readonly interpretationTone?: StatTone;
  /** Optional range/segment control shown beside the title. */
  readonly rangeControl?: ReactNode;
  /** The plot itself (a Victory chart or an insufficient-data message). */
  readonly children: ReactNode;
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
      return theme.color.textSecondary;
  }
}

/**
 * The mandatory chart wrapper (DESIGN_SYSTEM §6): title, the interpretation line, an
 * optional range control, and the plot. Charts never render outside a ChartFrame — a
 * plot without its coach sentence violates the analytics honesty contract.
 */
export function ChartFrame(props: ChartFrameProps): ReactNode {
  const theme = useTheme();
  return (
    <View style={{ gap: theme.space.sm }}>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: theme.space.sm,
        }}
      >
        <Text style={{ ...theme.type.heading, color: theme.color.textPrimary }}>{props.title}</Text>
        {props.rangeControl}
      </View>
      <Text
        style={{
          ...theme.type.caption,
          color: toneColor(theme, props.interpretationTone ?? 'neutral'),
        }}
      >
        {props.interpretation}
      </Text>
      <View style={{ marginTop: theme.space.xs }}>{props.children}</View>
    </View>
  );
}
