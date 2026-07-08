import { Minus, Plus } from 'lucide-react-native';
import { useEffect, useRef, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { triggerHaptic, useTheme } from '@/core/theme';

interface StepperProps {
  readonly value: number;
  readonly onChange: (next: number) => void;
  readonly step?: number;
  readonly min?: number;
  readonly max?: number;
  readonly format?: (value: number) => string;
  readonly accessibilityLabel?: string;
}

const REPEAT_DELAY_MS = 350;
const REPEAT_INTERVAL_MS = 120;

/** Numeric stepper — 44pt targets, long-press auto-repeat, haptic tick (DESIGN_SYSTEM §6). */
export function Stepper(props: StepperProps): ReactNode {
  const theme = useTheme();
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const valueRef = useRef(props.value);

  useEffect(() => {
    valueRef.current = props.value;
  }, [props.value]);

  useEffect(() => () => stopRepeat(timer), []);

  const step = props.step ?? 1;
  const apply = (direction: 1 | -1): void => {
    const raw = valueRef.current + direction * step;
    // Avoid float drift from repeated 0.1/2.5 steps.
    const next = Math.round(raw * 100) / 100;
    if (props.min !== undefined && next < props.min) return;
    if (props.max !== undefined && next > props.max) return;
    triggerHaptic('light');
    props.onChange(next);
  };
  const startRepeat = (direction: 1 | -1): void => {
    const tick = (): void => {
      apply(direction);
      timer.current = setTimeout(tick, REPEAT_INTERVAL_MS);
    };
    timer.current = setTimeout(tick, REPEAT_DELAY_MS);
  };

  const target = (direction: 1 | -1, icon: ReactNode, label: string): ReactNode => (
    <Pressable
      onPress={() => apply(direction)}
      onPressIn={() => startRepeat(direction)}
      onPressOut={() => stopRepeat(timer)}
      accessibilityRole="button"
      accessibilityLabel={label}
      style={({ pressed }) => ({
        width: 44,
        height: 44,
        alignItems: 'center',
        justifyContent: 'center',
        borderRadius: theme.radius.sm,
        backgroundColor: theme.color.surfaceRaised,
        opacity: pressed ? 0.7 : 1,
      })}
    >
      {icon}
    </Pressable>
  );

  return (
    <View
      style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}
      accessibilityLabel={props.accessibilityLabel}
    >
      {target(-1, <Minus color={theme.color.textPrimary} size={20} />, 'Decrease')}
      <Text
        style={{
          ...theme.type.bodyStrong,
          color: theme.color.textPrimary,
          fontVariant: ['tabular-nums'],
          minWidth: 64,
          textAlign: 'center',
        }}
      >
        {props.format ? props.format(props.value) : String(props.value)}
      </Text>
      {target(1, <Plus color={theme.color.textPrimary} size={20} />, 'Increase')}
    </View>
  );
}

function stopRepeat(timer: { current: ReturnType<typeof setTimeout> | null }): void {
  if (timer.current !== null) {
    clearTimeout(timer.current);
    timer.current = null;
  }
}
