import { Minus, Plus } from 'lucide-react-native';
import { useEffect, useRef, useState, type ReactNode } from 'react';
import { Pressable, Text, TextInput, View } from 'react-native';

import { triggerHaptic, useTheme } from '@/core/theme';

interface SetValueControlProps {
  readonly value: number;
  readonly onChange: (next: number) => void;
  readonly step: number;
  readonly max: number;
  readonly decimals: 0 | 1;
  readonly unit: string;
  readonly accessibilityLabel: string;
}

const REPEAT_DELAY_MS = 350;
const REPEAT_INTERVAL_MS = 110;

const roundTo = (value: number, decimals: 0 | 1): number => {
  const f = decimals === 1 ? 10 : 1;
  return Math.round(value * f) / f;
};

/**
 * The core logging control: large +/- targets for fast standing use, with the
 * value tappable to type (keyboard-first, UI_UX §5.3). Gym-optimized per Phase 4.
 */
export function SetValueControl(props: SetValueControlProps): ReactNode {
  const theme = useTheme();
  const [focused, setFocused] = useState(false);
  const [draft, setDraft] = useState('');
  const timer = useRef<ReturnType<typeof setInterval> | null>(null);
  const delay = useRef<ReturnType<typeof setTimeout> | null>(null);

  const stop = (): void => {
    if (delay.current) clearTimeout(delay.current);
    if (timer.current) clearInterval(timer.current);
    delay.current = null;
    timer.current = null;
  };

  useEffect(() => stop, []);

  const commit = (raw: number): void => {
    const next = Math.min(Math.max(roundTo(raw, props.decimals), 0), props.max);
    props.onChange(next);
  };
  const nudge = (dir: 1 | -1): void => {
    triggerHaptic('light');
    commit(props.value + dir * props.step);
  };
  const startRepeat = (dir: 1 | -1): void => {
    delay.current = setTimeout(() => {
      timer.current = setInterval(() => nudge(dir), REPEAT_INTERVAL_MS);
    }, REPEAT_DELAY_MS);
  };

  const display = focused ? draft : props.value.toFixed(props.decimals);

  const button = (dir: 1 | -1, icon: ReactNode, label: string): ReactNode => (
    <Pressable
      onPress={() => nudge(dir)}
      onPressIn={() => startRepeat(dir)}
      onPressOut={stop}
      accessibilityRole="button"
      accessibilityLabel={`${label} ${props.accessibilityLabel}`}
      // Visual 40px keeps a weight+reps row from overflowing (so the complete ✓ stays
      // on-screen); hitSlop restores a ≥44px touch target for standing gym use.
      hitSlop={theme.space.xs}
      style={({ pressed }) => ({
        width: 40,
        height: 44,
        alignItems: 'center',
        justifyContent: 'center',
        borderRadius: theme.radius.md,
        backgroundColor: theme.color.surfaceRaised,
        opacity: pressed ? 0.6 : 1,
      })}
    >
      {icon}
    </Pressable>
  );

  // flex:1 + minWidth:0 lets two controls share the row and shrink to fit, leaving
  // room for the complete ✓ instead of pushing it off the screen edge.
  return (
    <View style={{ flex: 1, minWidth: 0, alignItems: 'center', gap: theme.space.xs }}>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          alignSelf: 'stretch',
          justifyContent: 'center',
          gap: theme.space.xs,
        }}
      >
        {button(-1, <Minus color={theme.color.textPrimary} size={22} />, 'Decrease')}
        <TextInput
          value={display}
          onFocus={() => {
            setDraft(props.value.toFixed(props.decimals));
            setFocused(true);
          }}
          onChangeText={(text) => {
            setDraft(text);
            const parsed = Number.parseFloat(text);
            if (Number.isFinite(parsed)) commit(parsed);
          }}
          onBlur={() => setFocused(false)}
          keyboardType={props.decimals === 1 ? 'decimal-pad' : 'number-pad'}
          selectTextOnFocus
          accessibilityLabel={props.accessibilityLabel}
          style={{
            ...theme.type.heading,
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
            flex: 1,
            minWidth: 0,
            textAlign: 'center',
            paddingVertical: theme.space.xs,
          }}
        />
        {button(1, <Plus color={theme.color.textPrimary} size={22} />, 'Increase')}
      </View>
      <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>{props.unit}</Text>
    </View>
  );
}
