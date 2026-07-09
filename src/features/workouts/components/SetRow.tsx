import { Check } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { triggerHaptic, useTheme } from '@/core/theme';
import type { LoadType } from '@/domain/fitness';

import { useRestActions } from '../stores/useRestTimerStore';
import { useSessionActions, type SessionSet } from '../stores/useSessionStore';
import { PrBadge } from './PrBadge';
import { SetValueControl } from './SetValueControl';

interface SetRowProps {
  readonly exerciseLocalId: string;
  readonly index: number;
  readonly set: SessionSet;
  readonly loadType: LoadType;
  /** Optimistic: this set sets a new running weight/e1RM best (UI_UX §4.1, P15). */
  readonly isPr: boolean;
}

/** One logged set — big controls, warm-up toggle, and a one-tap complete (✓). */
export function SetRow(props: SetRowProps): ReactNode {
  const theme = useTheme();
  const actions = useSessionActions();
  const rest = useRestActions();
  const { set, loadType, exerciseLocalId, isPr } = props;

  const showWeight = loadType !== 'bodyweight' && loadType !== 'timed';
  const repsIsSeconds = loadType === 'timed';

  const complete = (): void => {
    const becomingDone = !set.done;
    // A PR completion gets the success haptic (delight registry); an ordinary set
    // gets light. Un-checking a set is silent.
    if (becomingDone) triggerHaptic(isPr ? 'success' : 'light');
    actions.toggleSetDone(exerciseLocalId, set.localId);
    // Auto-start rest only when a WORKING set is completed (UI_UX §4.2) — warm-ups
    // and un-checking a set never start a rest.
    if (becomingDone && !set.warmup) rest.start(exerciseLocalId, Date.now());
  };

  return (
    <View
      style={{
        flexDirection: 'row',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: theme.space.sm,
        paddingVertical: theme.space.sm,
        opacity: set.done ? 0.6 : 1,
      }}
    >
      <Pressable
        onPress={() => actions.updateSet(exerciseLocalId, set.localId, { warmup: !set.warmup })}
        accessibilityRole="button"
        accessibilityLabel={set.warmup ? 'Warm-up set' : 'Working set'}
        accessibilityState={{ selected: set.warmup }}
        style={{
          width: 28,
          height: 28,
          borderRadius: theme.radius.full,
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: set.warmup ? theme.color.accentSoft : 'transparent',
          borderWidth: 1,
          borderColor: set.warmup ? theme.color.accent : theme.color.border,
        }}
      >
        <Text
          style={{
            ...theme.type.micro,
            color: set.warmup ? theme.color.accent : theme.color.textSecondary,
          }}
        >
          {set.warmup ? 'W' : String(props.index + 1)}
        </Text>
      </Pressable>

      {showWeight ? (
        <SetValueControl
          value={set.weightKg}
          onChange={(next) => actions.updateSet(exerciseLocalId, set.localId, { weightKg: next })}
          step={2.5}
          max={1000}
          decimals={1}
          unit="KG"
          accessibilityLabel={`Set ${props.index + 1} weight`}
        />
      ) : null}

      <SetValueControl
        value={set.reps}
        onChange={(next) => actions.updateSet(exerciseLocalId, set.localId, { reps: next })}
        step={1}
        max={100}
        decimals={0}
        unit={repsIsSeconds ? 'SEC' : 'REPS'}
        accessibilityLabel={`Set ${props.index + 1} ${repsIsSeconds ? 'seconds' : 'reps'}`}
      />

      {set.done && isPr ? <PrBadge /> : null}

      <Pressable
        onPress={complete}
        accessibilityRole="button"
        accessibilityLabel={
          set.done ? `Set ${props.index + 1} done` : `Complete set ${props.index + 1}`
        }
        accessibilityState={{ checked: set.done }}
        style={{
          width: 48,
          height: 48,
          borderRadius: theme.radius.md,
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: set.done ? theme.color.positive : theme.color.surfaceRaised,
        }}
      >
        <Check
          color={set.done ? theme.color.bg : theme.color.textSecondary}
          size={24}
          strokeWidth={2.5}
        />
      </Pressable>
    </View>
  );
}
