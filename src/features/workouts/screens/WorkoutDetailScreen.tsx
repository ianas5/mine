import { useLocalSearchParams, useRouter } from 'expo-router';
import { ArrowLeft, Trash2 } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card, Dialog, EmptyState, IconButton, Screen, Skeleton, showToast } from '@/core/ui';
import { formatElapsed, formatRelativeDate } from '@/core/utils';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import type { Workout } from '@/domain/models';

import { DetailSetRow } from '../components/DetailSetRow';
import { useDefaultBodyweight } from '../hooks/useDefaultBodyweight';
import { useWorkout } from '../hooks/useWorkout';
import { workoutStats } from '../logic/workoutSummary';

/** Review and edit one past workout (Phase 5). Edits/deletes recompute everywhere. */
export function WorkoutDetailScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const params = useLocalSearchParams<{ id: string }>();
  const workout = useWorkout(params.id ?? '');

  return (
    <Screen scroll>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.sm,
          marginTop: theme.space.sm,
          marginBottom: theme.space.lg,
        }}
      >
        <IconButton
          icon={<ArrowLeft color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
          onPress={() => router.back()}
          accessibilityLabel="Back"
        />
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary, flex: 1 }}>
          {workout ? workout.name : 'Workout'}
        </Text>
      </View>

      {workout === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={72} />
          <Skeleton height={120} />
        </View>
      ) : workout === null ? (
        <EmptyState
          title="Workout not found"
          cta={{ label: 'Go back', onPress: () => router.back() }}
        />
      ) : (
        <LoadedDetail workout={workout} onDeleted={() => router.back()} />
      )}
    </Screen>
  );
}

function LoadedDetail(props: {
  readonly workout: Workout;
  readonly onDeleted: () => void;
}): ReactNode {
  const theme = useTheme();
  const bodyweightKg = useDefaultBodyweight();
  const [confirming, setConfirming] = useState(false);
  const { workout } = props;
  const stats = workoutStats(workout, bodyweightKg);
  const durationMs =
    workout.startedAt !== null && workout.endedAt !== null
      ? workout.endedAt - workout.startedAt
      : null;

  const stat = (value: string, label: string): ReactNode => (
    <View style={{ flex: 1, alignItems: 'center', gap: theme.space.xs }}>
      <Text
        style={{
          ...theme.type.heading,
          color: theme.color.textPrimary,
          fontVariant: ['tabular-nums'],
        }}
      >
        {value}
      </Text>
      <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>{label}</Text>
    </View>
  );

  return (
    <View style={{ gap: theme.space.lg }}>
      <View style={{ flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between' }}>
        <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
          {formatRelativeDate(workout.date)}
        </Text>
        <IconButton
          icon={<Trash2 color={theme.color.danger} size={22} strokeWidth={1.75} />}
          onPress={() => setConfirming(true)}
          accessibilityLabel="Delete workout"
        />
      </View>

      <Card style={{ flexDirection: 'row' }}>
        {stat(durationMs !== null ? formatElapsed(durationMs) : '—', 'DURATION')}
        {stat(String(stats.workingSetCount), 'WORKING SETS')}
        {stat(Math.round(stats.totalVolumeKg).toLocaleString(), 'VOLUME KG')}
      </Card>

      {workout.exercises.map((exercise) => (
        <Card key={exercise.id} style={{ gap: theme.space.xs }}>
          <Text style={{ ...theme.type.heading, color: theme.color.textPrimary }}>
            {exercise.name}
          </Text>
          {exercise.sets.map((set, index) => (
            <DetailSetRow key={set.id} set={set} index={index} loadType={exercise.loadType} />
          ))}
        </Card>
      ))}

      <Dialog
        visible={confirming}
        title="Delete workout?"
        message="This session and its sets will be permanently removed. Records may recede."
        confirmLabel="Delete"
        onConfirm={() => {
          setConfirming(false);
          void workoutRepository.remove(workout.id).then(() => {
            showToast('Workout deleted');
            props.onDeleted();
          });
        }}
        onCancel={() => setConfirming(false)}
      />
    </View>
  );
}
