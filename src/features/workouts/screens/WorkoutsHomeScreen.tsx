import { useRouter } from 'expo-router';
import { ChevronRight, Dumbbell, Play } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, Screen, Section } from '@/core/ui';

import { RecentWorkoutList } from '../components/RecentWorkoutList';
import { useSessionActions, useSessionStore } from '../stores/useSessionStore';

/** Workouts home — start a session or manage the catalog (grows in Phases 5/8). */
export function WorkoutsHomeScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const actions = useSessionActions();
  const sessionActive = useSessionStore((s) => s.active);

  const startEmpty = (): void => {
    actions.start(Date.now(), 'Workout');
    router.push('/active-workout');
  };

  return (
    <Screen scroll>
      <Text
        style={{
          ...theme.type.title,
          color: theme.color.textPrimary,
          marginTop: theme.space.sm,
          marginBottom: theme.space.xl,
        }}
      >
        Workouts
      </Text>

      {sessionActive ? (
        <Card variant="accentEdge" style={{ marginBottom: theme.space.lg }}>
          <Text style={{ ...theme.type.heading, color: theme.color.textPrimary }}>
            Workout in progress
          </Text>
          <Text
            style={{
              ...theme.type.caption,
              color: theme.color.textSecondary,
              marginTop: theme.space.xs,
              marginBottom: theme.space.md,
            }}
          >
            Pick up where you left off.
          </Text>
          <Button label="Resume workout" onPress={() => router.push('/active-workout')} />
        </Card>
      ) : (
        <View style={{ marginBottom: theme.space.lg, gap: theme.space.sm }}>
          <Button label="Start empty workout" onPress={startEmpty} />
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
              Programs & templates arrive in a later update
            </Text>
          </View>
        </View>
      )}

      <Pressable
        onPress={() => router.push('/workouts/library')}
        accessibilityRole="button"
        accessibilityLabel="Manage exercises"
        style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
      >
        <Card
          style={{
            flexDirection: 'row',
            alignItems: 'center',
            gap: theme.space.md,
          }}
        >
          <Dumbbell color={theme.color.accent} size={22} strokeWidth={1.75} />
          <View style={{ flex: 1 }}>
            <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>Exercises</Text>
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              Browse, search, and add custom exercises
            </Text>
          </View>
          <ChevronRight color={theme.color.textTertiary} size={20} />
        </Card>
      </Pressable>

      <View style={{ marginTop: theme.space.xl }}>
        <Section title="Recent">
          <RecentWorkoutList />
        </Section>
      </View>
    </Screen>
  );
}
