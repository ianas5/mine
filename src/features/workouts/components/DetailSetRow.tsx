import { Trash2 } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { workoutRepository } from '@/data/repositories/workoutRepository';
import type { LoadType } from '@/domain/fitness';
import type { WorkoutSet } from '@/domain/models';

import { SetValueControl } from './SetValueControl';

interface DetailSetRowProps {
  readonly set: WorkoutSet;
  readonly index: number;
  readonly loadType: LoadType;
}

/** An editable saved set on the workout-detail screen; edits recompute via change-bus. */
export function DetailSetRow(props: DetailSetRowProps): ReactNode {
  const theme = useTheme();
  const { set, loadType } = props;
  const [weightKg, setWeightKg] = useState(set.weightKg);
  const [reps, setReps] = useState(set.reps);
  const [warmup, setWarmup] = useState(set.warmup);

  const showWeight = loadType !== 'bodyweight' && loadType !== 'timed';

  return (
    <View
      style={{
        flexDirection: 'row',
        alignItems: 'center',
        gap: theme.space.xs,
        paddingVertical: theme.space.sm,
      }}
    >
      <Pressable
        onPress={() => {
          const next = !warmup;
          setWarmup(next);
          void workoutRepository.updateSet(set.id, { warmup: next });
        }}
        accessibilityRole="button"
        accessibilityLabel={warmup ? 'Warm-up set' : 'Working set'}
        accessibilityState={{ selected: warmup }}
        style={{
          width: 28,
          height: 28,
          flexShrink: 0,
          borderRadius: theme.radius.full,
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: warmup ? theme.color.accentSoft : 'transparent',
          borderWidth: 1,
          borderColor: warmup ? theme.color.accent : theme.color.border,
        }}
      >
        <Text
          style={{
            ...theme.type.micro,
            color: warmup ? theme.color.accent : theme.color.textSecondary,
          }}
        >
          {warmup ? 'W' : String(props.index + 1)}
        </Text>
      </Pressable>

      {showWeight ? (
        <SetValueControl
          value={weightKg}
          onChange={(next) => {
            setWeightKg(next);
            void workoutRepository.updateSet(set.id, { weightKg: next });
          }}
          step={2.5}
          max={1000}
          decimals={1}
          unit="KG"
          accessibilityLabel={`Set ${props.index + 1} weight`}
        />
      ) : null}

      <SetValueControl
        value={reps}
        onChange={(next) => {
          setReps(next);
          void workoutRepository.updateSet(set.id, { reps: next });
        }}
        step={1}
        max={100}
        decimals={0}
        unit={loadType === 'timed' ? 'SEC' : 'REPS'}
        accessibilityLabel={`Set ${props.index + 1} reps`}
      />

      <Pressable
        onPress={() => void workoutRepository.deleteSet(set.id)}
        accessibilityRole="button"
        accessibilityLabel={`Delete set ${props.index + 1}`}
        hitSlop={theme.space.sm}
        style={{
          width: 44,
          height: 44,
          flexShrink: 0,
          alignItems: 'center',
          justifyContent: 'center',
        }}
      >
        <Trash2 color={theme.color.textTertiary} size={20} />
      </Pressable>
    </View>
  );
}
