import { Sparkles, TrendingUp, TriangleAlert, X } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';

export type InsightCardTone = 'positive' | 'attention' | 'neutral';

interface InsightCardProps {
  readonly tone: InsightCardTone;
  readonly title: string;
  readonly body: string;
  /** Tap-through to the insight's evidence (UI_UX §8). */
  readonly onPress?: () => void;
  readonly onDismiss?: () => void;
}

function toneColor(theme: Theme, tone: InsightCardTone): string {
  if (tone === 'positive') return theme.color.positive;
  if (tone === 'attention') return theme.color.attention;
  return theme.color.textSecondary;
}

function ToneIcon(props: { readonly tone: InsightCardTone; readonly color: string }): ReactNode {
  if (props.tone === 'positive') return <TrendingUp color={props.color} size={18} />;
  if (props.tone === 'attention') return <TriangleAlert color={props.color} size={18} />;
  return <Sparkles color={props.color} size={18} />;
}

/**
 * A generated coach insight (DESIGN_SYSTEM §6): tone-tinted left edge + icon, a heading,
 * a 1–2 sentence body, an optional evidence tap-through, and a dismiss control. Purely
 * presentational — wording and routing are decided by the engine/feature.
 */
export function InsightCard(props: InsightCardProps): ReactNode {
  const theme = useTheme();
  const accent = toneColor(theme, props.tone);

  return (
    <Pressable
      onPress={props.onPress}
      accessibilityRole={props.onPress ? 'button' : undefined}
      accessibilityLabel={props.title}
      style={({ pressed }) => ({
        flexDirection: 'row',
        gap: theme.space.md,
        backgroundColor: theme.color.surface,
        borderRadius: theme.radius.md,
        borderWidth: 1,
        borderColor: theme.color.border,
        borderLeftWidth: 3,
        borderLeftColor: accent,
        padding: theme.space.lg,
        opacity: pressed && props.onPress ? 0.85 : 1,
      })}
    >
      <View style={{ paddingTop: 2 }}>
        <ToneIcon tone={props.tone} color={accent} />
      </View>
      <View style={{ flex: 1, gap: theme.space.xs }}>
        <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
          {props.title}
        </Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
          {props.body}
        </Text>
      </View>
      {props.onDismiss ? (
        <Pressable
          onPress={props.onDismiss}
          accessibilityRole="button"
          accessibilityLabel="Dismiss"
          hitSlop={8}
          style={({ pressed }) => ({ opacity: pressed ? 0.5 : 1 })}
        >
          <X color={theme.color.textTertiary} size={18} />
        </Pressable>
      ) : null}
    </Pressable>
  );
}
