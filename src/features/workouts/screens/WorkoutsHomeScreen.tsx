import { useRouter } from 'expo-router';
import { CalendarRange, ChevronRight, Dumbbell } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, Dialog, Screen, Section } from '@/core/ui';

import { RecentWorkoutList } from '../components/RecentWorkoutList';
import { StartSection } from '../components/StartSection';
import { useRestActions } from '../stores/useRestTimerStore';
import { useSessionActions, useSessionStore } from '../stores/useSessionStore';

/** Workouts home — start a session or manage the catalog (grows in Phases 5/8). */
export function WorkoutsHomeScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const actions = useSessionActions();
  const restActions = useRestActions();
  const sessionActive = useSessionStore((s) => s.active);
  const recovered = useSessionStore((s) => s.recovered);
  const [confirmDiscard, setConfirmDiscard] = useState(false);

  const resume = (): void => {
    actions.acknowledgeRecovery();
    router.push('/active-workout');
  };

  const discardRecovered = (): void => {
    setConfirmDiscard(false);
    restActions.reset();
    actions.discard();
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
            {recovered ? 'Workout recovered' : 'Workout in progress'}
          </Text>
          <Text
            style={{
              ...theme.type.caption,
              color: theme.color.textSecondary,
              marginTop: theme.space.xs,
              marginBottom: theme.space.md,
            }}
          >
            {recovered
              ? 'Your session was saved and restored. Pick up right where you left off.'
              : 'Pick up where you left off.'}
          </Text>
          <Button label="Resume workout" onPress={resume} />
          {recovered ? (
            <Pressable
              onPress={() => setConfirmDiscard(true)}
              accessibilityRole="button"
              accessibilityLabel="Discard recovered workout"
              style={({ pressed }) => ({
                alignSelf: 'center',
                paddingVertical: theme.space.sm,
                marginTop: theme.space.xs,
                opacity: pressed ? 0.6 : 1,
              })}
            >
              <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>
                Discard
              </Text>
            </Pressable>
          ) : null}
        </Card>
      ) : (
        <StartSection />
      )}

      <View style={{ gap: theme.space.sm }}>
        <NavRow
          icon={<CalendarRange color={theme.color.accent} size={22} strokeWidth={1.75} />}
          title="Programs"
          subtitle="Plan sessions, weekdays, and targets"
          onPress={() => router.push('/workouts/programs')}
        />
        <NavRow
          icon={<Dumbbell color={theme.color.accent} size={22} strokeWidth={1.75} />}
          title="Exercises"
          subtitle="Browse, search, and add custom exercises"
          onPress={() => router.push('/workouts/library')}
        />
      </View>

      <View style={{ marginTop: theme.space.xl }}>
        <Section title="Recent">
          <RecentWorkoutList />
        </Section>
      </View>

      <Dialog
        visible={confirmDiscard}
        title="Discard recovered workout?"
        message="The restored session and its logged sets will be permanently removed."
        confirmLabel="Discard"
        onConfirm={discardRecovered}
        onCancel={() => setConfirmDiscard(false)}
      />
    </Screen>
  );
}

function NavRow(props: {
  readonly icon: ReactNode;
  readonly title: string;
  readonly subtitle: string;
  readonly onPress: () => void;
}): ReactNode {
  const theme = useTheme();
  return (
    <Pressable
      onPress={props.onPress}
      accessibilityRole="button"
      accessibilityLabel={props.title}
      style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
    >
      <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
        {props.icon}
        <View style={{ flex: 1 }}>
          <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>{props.title}</Text>
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            {props.subtitle}
          </Text>
        </View>
        <ChevronRight color={theme.color.textTertiary} size={20} />
      </Card>
    </Pressable>
  );
}
