import type { ReactNode } from 'react';
import { Pressable, Text } from 'react-native';

import { triggerHaptic, useTheme } from '@/core/theme';

interface ChipProps {
  readonly label: string;
  readonly selected: boolean;
  readonly onPress: () => void;
}

/** Selectable pill — accentSoft when selected; haptic light on select. */
export function Chip(props: ChipProps): ReactNode {
  const theme = useTheme();
  return (
    <Pressable
      onPress={() => {
        if (!props.selected) {
          triggerHaptic('light');
        }
        props.onPress();
      }}
      accessibilityRole="button"
      accessibilityLabel={props.label}
      accessibilityState={{ selected: props.selected }}
      style={({ pressed }) => ({
        paddingHorizontal: theme.space.lg,
        paddingVertical: theme.space.sm,
        borderRadius: theme.radius.full,
        backgroundColor: props.selected ? theme.color.accentSoft : theme.color.surface,
        borderWidth: 1,
        borderColor: props.selected ? theme.color.accent : theme.color.border,
        opacity: pressed ? 0.8 : 1,
      })}
    >
      <Text
        style={{
          ...theme.type.caption,
          color: props.selected ? theme.color.accent : theme.color.textSecondary,
        }}
      >
        {props.label}
      </Text>
    </Pressable>
  );
}
