import { useMemo, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Sheet } from '@/core/ui';
import { formatElapsed } from '@/core/utils';
import { computeWorkoutStats } from '@/domain/fitness';

import { sessionToStatExercises } from '../logic/sessionMapping';
import type { SessionExercise } from '../stores/useSessionStore';

interface WorkoutSummarySheetProps {
  readonly visible: boolean;
  readonly exercises: readonly SessionExercise[];
  readonly bodyweightKg: number | null;
  readonly elapsedMs: number;
  readonly prCount: number;
  readonly saving: boolean;
  readonly onSave: () => void;
  readonly onDiscard: () => void;
  readonly onClose: () => void;
}

/** Finish summary — duration, volume, working sets, and PRs earned (FITNESS_DOMAIN §3.7). */
export function WorkoutSummarySheet(props: WorkoutSummarySheetProps): ReactNode {
  const theme = useTheme();
  const stats = useMemo(
    () => computeWorkoutStats(sessionToStatExercises(props.exercises), props.bodyweightKg),
    [props.exercises, props.bodyweightKg],
  );

  const stat = (value: string, label: string): ReactNode => (
    <View style={{ flex: 1, alignItems: 'center', gap: theme.space.xs }}>
      <Text
        style={{
          ...theme.type.display,
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
    <Sheet visible={props.visible} onClose={props.onClose} title="Finish Workout">
      <View style={{ gap: theme.space.xl }}>
        <View style={{ flexDirection: 'row' }}>
          {stat(formatElapsed(props.elapsedMs), 'DURATION')}
          {stat(String(stats.workingSetCount), 'WORKING SETS')}
          {stat(`${Math.round(stats.totalVolumeKg).toLocaleString()}`, 'VOLUME KG')}
        </View>
        {props.prCount > 0 ? (
          <View
            style={{
              alignItems: 'center',
              paddingVertical: theme.space.md,
              borderRadius: theme.radius.md,
              backgroundColor: theme.color.accentSoft,
            }}
          >
            <Text style={{ ...theme.type.heading, color: theme.color.accent }}>
              {props.prCount} new {props.prCount === 1 ? 'PR' : 'PRs'}
            </Text>
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              A new personal record this session
            </Text>
          </View>
        ) : null}
        {stats.volumeLowConfidence ? (
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            Some volume is estimated — set a default bodyweight in Settings for bodyweight
            exercises.
          </Text>
        ) : null}
        <Button label="Save workout" onPress={props.onSave} loading={props.saving} />
        <Button label="Discard workout" variant="ghost" size="md" onPress={props.onDiscard} />
      </View>
    </Sheet>
  );
}
