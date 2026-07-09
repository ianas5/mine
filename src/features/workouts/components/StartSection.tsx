import { Play, Repeat } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card } from '@/core/ui';
import { weekdayLabel } from '@/core/utils';

import { useRecentWorkouts } from '../hooks/useRecentWorkouts';
import { useStartWorkout } from '../hooks/useStartWorkout';
import { useTemplateSuggestion } from '../hooks/useTemplateSuggestion';

/**
 * The start-a-workout surface (UI_UX §5.2): a one-tap smart-default start (the
 * active program's weekday template → most-frequent → repeat last), plus an
 * always-available empty start. Starting from the suggestion is ≤ 1 tap.
 */
export function StartSection(): ReactNode {
  const theme = useTheme();
  const suggestion = useTemplateSuggestion();
  const recent = useRecentWorkouts();
  const lastWorkout = recent?.[0] ?? null;
  const { startEmpty, startFromTemplate, startRepeatLast } = useStartWorkout();

  const suggestionCard = (title: string, subtitle: string, onPress: () => void): ReactNode => (
    <Pressable
      onPress={onPress}
      accessibilityRole="button"
      accessibilityLabel={title}
      style={({ pressed }) => ({ opacity: pressed ? 0.85 : 1 })}
    >
      <Card variant="accentEdge" style={{ gap: theme.space.xs }}>
        <Text style={{ ...theme.type.micro, color: theme.color.accent }}>SUGGESTED FOR TODAY</Text>
        <Text style={{ ...theme.type.heading, color: theme.color.textPrimary }}>{title}</Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>{subtitle}</Text>
      </Card>
    </Pressable>
  );

  return (
    <View style={{ marginBottom: theme.space.lg, gap: theme.space.sm }}>
      {suggestion?.kind === 'template'
        ? suggestionCard(
            suggestion.template.name,
            suggestion.template.weekday !== null
              ? `Scheduled for ${weekdayLabel(suggestion.template.weekday)} · tap to start`
              : 'Tap to start with your targets',
            () => void startFromTemplate(suggestion.template),
          )
        : null}
      {suggestion?.kind === 'repeatLast'
        ? suggestionCard(`Repeat ${suggestion.workout.name}`, 'Same exercises, ready to log', () =>
            startRepeatLast(suggestion.workout),
          )
        : null}

      <Button label="Start empty workout" onPress={startEmpty} />

      {suggestion?.kind === 'template' && lastWorkout ? (
        <Pressable
          onPress={() => startRepeatLast(lastWorkout)}
          accessibilityRole="button"
          accessibilityLabel="Repeat last workout"
          style={({ pressed }) => ({
            flexDirection: 'row',
            alignItems: 'center',
            justifyContent: 'center',
            gap: theme.space.xs,
            paddingVertical: theme.space.sm,
            opacity: pressed ? 0.6 : 1,
          })}
        >
          <Repeat color={theme.color.textTertiary} size={14} />
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            Repeat last workout
          </Text>
        </Pressable>
      ) : null}

      {suggestion?.kind === 'none' ? (
        <View
          style={{
            flexDirection: 'row',
            alignItems: 'center',
            justifyContent: 'center',
            gap: theme.space.xs,
          }}
        >
          <Play color={theme.color.textTertiary} size={14} />
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            Build a program below to get weekday suggestions
          </Text>
        </View>
      ) : null}
    </View>
  );
}
