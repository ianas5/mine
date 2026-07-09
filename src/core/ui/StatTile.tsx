import type { ReactNode } from 'react';
import { Text, View, type TextStyle } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';

/** Interpretation classification → context-line colour (ANALYTICS §2.2, §5.3). */
export type StatTone = 'positive' | 'negative' | 'neutral' | 'attention';

interface StatTileProps {
  readonly label: string;
  /** The value, pre-rounded and pre-formatted by the caller (never raw math here). */
  readonly value: string;
  readonly unit?: string;
  /**
   * The **required** context line — reference + classification (e.g. "↓ 1.2 kg vs
   * last month · improving"). Enforces the interpretation triplet (ANALYTICS rule 3):
   * a StatTile literally cannot be rendered without it.
   */
  readonly context: string;
  readonly tone?: StatTone;
  /** Optional leading glyph for the context line (a trend arrow icon). */
  readonly glyph?: ReactNode;
}

function toneColor(theme: Theme, tone: StatTone): string {
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
 * The interpretation-triplet primitive (DESIGN_SYSTEM §6, P6/P8): value + unit + label
 * + a mandatory context line carrying the reference and classification. There is no
 * context-less variant by design — a number with no interpretation is a rendering bug.
 */
export function StatTile(props: StatTileProps): ReactNode {
  const theme = useTheme();
  const color = toneColor(theme, props.tone ?? 'neutral');
  return (
    <View style={{ gap: theme.space.xs }}>
      <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>
        {props.label.toUpperCase()}
      </Text>
      <View style={{ flexDirection: 'row', alignItems: 'baseline', gap: theme.space.xs }}>
        <Text
          style={{
            ...(theme.type.title as TextStyle),
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {props.value}
        </Text>
        {props.unit ? (
          <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>
            {props.unit}
          </Text>
        ) : null}
      </View>
      <View style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.xs }}>
        {props.glyph}
        <Text style={{ ...theme.type.caption, color }}>{props.context}</Text>
      </View>
    </View>
  );
}
