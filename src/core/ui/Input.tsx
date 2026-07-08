import { useState, type ReactNode } from 'react';
import { Text, TextInput, View, type KeyboardTypeOptions, type TextStyle } from 'react-native';

import { useTheme } from '@/core/theme';

interface InputProps {
  readonly value: string;
  readonly onChangeText: (text: string) => void;
  readonly label?: string;
  readonly error?: string;
  readonly placeholder?: string;
  readonly unit?: string;
  readonly keyboardType?: KeyboardTypeOptions;
  readonly returnKeyType?: 'next' | 'done';
  readonly onSubmitEditing?: () => void;
  readonly onBlur?: () => void;
  readonly autoFocus?: boolean;
  readonly editable?: boolean;
  readonly accessibilityLabel?: string;
}

/** Text/numeric input — RHF-compatible props; numeric values select-on-focus (UI_UX §5.3). */
export function Input(props: InputProps): ReactNode {
  const theme = useTheme();
  const [focused, setFocused] = useState(false);
  const isNumeric = props.keyboardType === 'decimal-pad' || props.keyboardType === 'number-pad';
  const borderColor =
    props.error !== undefined
      ? theme.color.danger
      : focused
        ? theme.color.accent
        : theme.color.border;

  const valueStyle: TextStyle = {
    ...(theme.type.body as TextStyle),
    color: theme.color.textPrimary,
    flex: 1,
    paddingVertical: theme.space.md,
    ...(isNumeric && { fontVariant: ['tabular-nums'], textAlign: 'right' as const }),
  };

  return (
    <View style={{ gap: theme.space.xs }}>
      {props.label !== undefined ? (
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
          {props.label}
        </Text>
      ) : null}
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          borderWidth: 1,
          borderColor,
          borderRadius: theme.radius.sm,
          paddingHorizontal: theme.space.md,
          backgroundColor: theme.color.surface,
          gap: theme.space.sm,
        }}
      >
        <TextInput
          value={props.value}
          onChangeText={props.onChangeText}
          placeholder={props.placeholder}
          placeholderTextColor={theme.color.textTertiary}
          keyboardType={props.keyboardType}
          returnKeyType={props.returnKeyType}
          onSubmitEditing={props.onSubmitEditing}
          autoFocus={props.autoFocus}
          editable={props.editable}
          selectTextOnFocus={isNumeric}
          onFocus={() => setFocused(true)}
          onBlur={() => {
            setFocused(false);
            props.onBlur?.();
          }}
          accessibilityLabel={props.accessibilityLabel ?? props.label}
          style={valueStyle}
        />
        {props.unit !== undefined ? (
          <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
            {props.unit}
          </Text>
        ) : null}
      </View>
      {props.error !== undefined ? (
        <Text
          style={{ ...theme.type.caption, color: theme.color.danger }}
          accessibilityRole="alert"
        >
          {props.error}
        </Text>
      ) : null}
    </View>
  );
}
