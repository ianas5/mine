import type { ReactNode } from 'react';
import { ActivityIndicator, Pressable, Text, View, type TextStyle } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';

type ButtonVariant = 'primary' | 'secondary' | 'ghost' | 'destructive';
type ButtonSize = 'lg' | 'md';

interface ButtonProps {
  readonly label: string;
  readonly onPress: () => void;
  readonly variant?: ButtonVariant;
  readonly size?: ButtonSize;
  readonly disabled?: boolean;
  readonly loading?: boolean;
}

/** Button — variants/sizes/states per DESIGN_SYSTEM §6; loading locks width. */
export function Button(props: ButtonProps): ReactNode {
  const theme = useTheme();
  const variant = props.variant ?? 'primary';
  const size = props.size ?? 'lg';
  const disabled = props.disabled === true || props.loading === true;
  const colors = variantColors(variant, theme);

  return (
    <Pressable
      onPress={props.onPress}
      disabled={disabled}
      accessibilityRole="button"
      accessibilityLabel={props.label}
      accessibilityState={{ disabled, busy: props.loading === true }}
      style={({ pressed }) => ({
        height: size === 'lg' ? 52 : 44,
        borderRadius: theme.radius.md,
        alignItems: 'center',
        justifyContent: 'center',
        paddingHorizontal: theme.space.xl,
        backgroundColor: colors.bg,
        borderWidth: variant === 'secondary' ? 1 : 0,
        borderColor: theme.color.border,
        opacity: props.disabled === true ? 0.4 : 1,
        transform: [{ scale: pressed && !disabled ? 0.98 : 1 }],
      })}
    >
      <Text
        style={{
          ...(theme.type.bodyStrong as TextStyle),
          color: colors.fg,
          opacity: props.loading === true ? 0 : 1,
        }}
      >
        {props.label}
      </Text>
      {props.loading === true ? (
        <View style={{ position: 'absolute' }}>
          <ActivityIndicator color={colors.fg} accessibilityLabel="Loading" />
        </View>
      ) : null}
    </Pressable>
  );
}

function variantColors(variant: ButtonVariant, theme: Theme): { bg: string; fg: string } {
  switch (variant) {
    case 'primary':
      return { bg: theme.color.accent, fg: theme.color.textPrimary };
    case 'secondary':
      return { bg: theme.color.surface, fg: theme.color.textPrimary };
    case 'ghost':
      return { bg: 'transparent', fg: theme.color.accent };
    case 'destructive':
      return { bg: theme.color.danger, fg: theme.color.textPrimary };
  }
}
