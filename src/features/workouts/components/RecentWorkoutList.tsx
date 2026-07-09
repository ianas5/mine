import { useRouter } from 'expo-router';
import { ChevronRight } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card, EmptyState, Skeleton } from '@/core/ui';
import { formatRelativeDate } from '@/core/utils';
import type { Workout } from '@/domain/models';

import { useDefaultBodyweight } from '../hooks/useDefaultBodyweight';
import { useRecentWorkouts } from '../hooks/useRecentWorkouts';
import { workoutStats } from '../logic/workoutSummary';

/** Recent workouts on the Workouts home — tap through to the detail/editor. */
export function RecentWorkoutList(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const workouts = useRecentWorkouts();
  const bodyweightKg = useDefaultBodyweight();

  if (workouts === null) {
    return (
      <View style={{ gap: theme.space.sm }}>
        <Skeleton height={64} />
        <Skeleton height={64} />
      </View>
    );
  }
  if (workouts.length === 0) {
    return <EmptyState title="No workouts yet — start one above" />;
  }

  return (
    <View style={{ gap: theme.space.sm }}>
      {workouts.map((workout) => (
        <WorkoutRow
          key={workout.id}
          workout={workout}
          bodyweightKg={bodyweightKg}
          onOpen={() => router.push(`/workouts/${workout.id}`)}
        />
      ))}
    </View>
  );
}

function WorkoutRow(props: {
  readonly workout: Workout;
  readonly bodyweightKg: number | null;
  readonly onOpen: () => void;
}): ReactNode {
  const theme = useTheme();
  const stats = workoutStats(props.workout, props.bodyweightKg);
  return (
    <Pressable
      onPress={props.onOpen}
      accessibilityRole="button"
      accessibilityLabel={`${props.workout.name}, ${formatRelativeDate(props.workout.date)}`}
      style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
    >
      <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
        <View style={{ flex: 1 }}>
          <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
            {props.workout.name}
          </Text>
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            {formatRelativeDate(props.workout.date)} · {props.workout.exercises.length} exercises ·{' '}
            {stats.workingSetCount} sets · {Math.round(stats.totalVolumeKg).toLocaleString()} kg
          </Text>
        </View>
        <ChevronRight color={theme.color.textTertiary} size={20} />
      </Card>
    </Pressable>
  );
}
