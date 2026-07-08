import type { ReactNode } from 'react';
import { Pressable } from 'react-native';

import { useTheme } from '@/core/theme';

interface IconButtonProps {
  readonly icon: ReactNode;
  readonly onPress: () => void;
  /** Required — icons alone carry no text (DESIGN_SYSTEM §8). */
  readonly accessibilityLabel: string;
  readonly disabled?: boolean;
}

/** Icon-only button — 44pt minimum target regardless of glyph size. */
export function IconButton(props: IconButtonProps): ReactNode {
  const theme = useTheme();
  return (
    <Pressable
      onPress={props.onPress}
      disabled={props.disabled === true}
      accessibilityRole="button"
      accessibilityLabel={props.accessibilityLabel}
      accessibilityState={{ disabled: props.disabled === true }}
      style={({ pressed }) => ({
        minWidth: 44,
        minHeight: 44,
        alignItems: 'center',
        justifyContent: 'center',
        borderRadius: theme.radius.md,
        opacity: props.disabled === true ? 0.4 : pressed ? 0.7 : 1,
      })}
    >
      {props.icon}
    </Pressable>
  );
}
