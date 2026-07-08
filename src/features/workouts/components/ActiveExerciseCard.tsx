import { Plus, X } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card, Chip } from '@/core/ui';

import { useSessionActions, type SessionExercise } from '../stores/useSessionStore';
import { SetRow } from './SetRow';

interface ActiveExerciseCardProps {
  readonly exercise: SessionExercise;
}

/** One exercise within the active session: header, set rows, and Add set. */
export function ActiveExerciseCard(props: ActiveExerciseCardProps): ReactNode {
  const theme = useTheme();
  const actions = useSessionActions();
  const { exercise } = props;
  const isUnilateral = exercise.unilateralCounting !== 'none';

  return (
    <Card style={{ gap: theme.space.sm }}>
      <View style={{ flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between' }}>
        <Text style={{ ...theme.type.heading, color: theme.color.textPrimary, flex: 1 }}>
          {exercise.name}
        </Text>
        <Pressable
          onPress={() => actions.removeExercise(exercise.localId)}
          accessibilityRole="button"
          accessibilityLabel={`Remove ${exercise.name}`}
          hitSlop={theme.space.sm}
          style={{ padding: theme.space.xs }}
        >
          <X color={theme.color.textTertiary} size={20} />
        </Pressable>
      </View>

      {isUnilateral ? (
        <View style={{ flexDirection: 'row', gap: theme.space.sm }}>
          <Chip
            label="×2 (one side)"
            selected={exercise.unilateralCounting === 'single_doubled'}
            onPress={() => actions.setCounting(exercise.localId, 'single_doubled')}
          />
          <Chip
            label="Per side logged"
            selected={exercise.unilateralCounting === 'per_side'}
            onPress={() => actions.setCounting(exercise.localId, 'per_side')}
          />
        </View>
      ) : null}

      <View>
        {exercise.sets.map((set, index) => (
          <SetRow
            key={set.localId}
            exerciseLocalId={exercise.localId}
            index={index}
            set={set}
            loadType={exercise.loadType}
          />
        ))}
      </View>

      <Pressable
        onPress={() => actions.addSet(exercise.localId)}
        accessibilityRole="button"
        accessibilityLabel={`Add set to ${exercise.name}`}
        style={({ pressed }) => ({
          flexDirection: 'row',
          alignItems: 'center',
          justifyContent: 'center',
          gap: theme.space.xs,
          paddingVertical: theme.space.md,
          borderRadius: theme.radius.md,
          borderWidth: 1,
          borderColor: theme.color.border,
          borderStyle: 'dashed',
          opacity: pressed ? 0.6 : 1,
        })}
      >
        <Plus color={theme.color.accent} size={18} />
        <Text style={{ ...theme.type.bodyStrong, color: theme.color.accent }}>Add set</Text>
      </Pressable>
    </Card>
  );
}
