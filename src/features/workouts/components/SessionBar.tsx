import { useRouter } from 'expo-router';
import { ChevronUp, Timer } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { formatElapsed } from '@/core/utils';

import { useElapsed } from '../hooks/useElapsed';
import { formatCountdown, useRestCountdown } from '../hooks/useRestCountdown';
import { useSessionStore, type SessionExercise } from '../stores/useSessionStore';

/** The exercise the lifter is on: first with an unfinished set, else the last one. */
function currentExerciseName(exercises: readonly SessionExercise[]): string | null {
  const inProgress = exercises.find((ex) => ex.sets.some((s) => !s.done));
  const exercise = inProgress ?? exercises[exercises.length - 1];
  return exercise ? exercise.name : null;
}

/**
 * Persistent session bar docked above the tab bar app-wide (UI_UX §2.2/§5.1 Focus
 * Mode): elapsed · current exercise · rest countdown when running. One tap returns
 * to the live session. Renders nothing when no workout is active.
 */
export function SessionBar(): ReactNode {
  const theme = useTheme();
  const router = useRouter();

  const active = useSessionStore((s) => s.active);
  const startedAt = useSessionStore((s) => s.startedAt);
  const exercises = useSessionStore((s) => s.exercises);
  const elapsedMs = useElapsed(startedAt);
  const { running: resting, remainingMs } = useRestCountdown();

  if (!active) return null;

  const current = currentExerciseName(exercises);

  return (
    <Pressable
      onPress={() => router.push('/active-workout')}
      accessibilityRole="button"
      accessibilityLabel="Return to active workout"
      style={({ pressed }) => ({
        flexDirection: 'row',
        alignItems: 'center',
        gap: theme.space.sm,
        paddingHorizontal: theme.space.lg,
        paddingVertical: theme.space.sm,
        backgroundColor: theme.color.surfaceRaised,
        borderTopWidth: 1,
        borderTopColor: theme.color.border,
        opacity: pressed ? 0.85 : 1,
      })}
    >
      <View
        style={{
          width: 8,
          height: 8,
          borderRadius: theme.radius.full,
          backgroundColor: theme.color.accent,
        }}
      />
      <View style={{ flex: 1 }}>
        <Text
          style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}
          numberOfLines={1}
        >
          {current ?? 'Workout in progress'}
        </Text>
        <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>Tap to return</Text>
      </View>

      {resting ? (
        <View style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.xs }}>
          <Timer color={theme.color.accent} size={16} strokeWidth={2} />
          <Text
            style={{
              ...theme.type.bodyStrong,
              color: theme.color.accent,
              fontVariant: ['tabular-nums'],
            }}
          >
            {formatCountdown(remainingMs)}
          </Text>
        </View>
      ) : (
        <Text
          style={{
            ...theme.type.bodyStrong,
            color: theme.color.textSecondary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {formatElapsed(elapsedMs)}
        </Text>
      )}

      <ChevronUp color={theme.color.textTertiary} size={20} />
    </Pressable>
  );
}
