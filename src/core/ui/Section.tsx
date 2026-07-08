import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';

interface SectionProps {
  readonly title: string;
  readonly action?: { readonly label: string; readonly onPress: () => void };
  readonly children: ReactNode;
}

/** Section header + content gap — the only way section headers are built. */
export function Section(props: SectionProps): ReactNode {
  const theme = useTheme();
  return (
    <View style={{ marginBottom: theme.space.xxl }}>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          justifyContent: 'space-between',
          marginBottom: theme.space.md,
        }}
      >
        <Text style={{ ...theme.type.heading, color: theme.color.textPrimary }}>{props.title}</Text>
        {props.action ? (
          <Pressable
            onPress={props.action.onPress}
            accessibilityRole="button"
            accessibilityLabel={props.action.label}
            hitSlop={theme.space.sm}
          >
            <Text style={{ ...theme.type.caption, color: theme.color.accent }}>
              {props.action.label}
            </Text>
          </Pressable>
        ) : null}
      </View>
      {props.children}
    </View>
  );
}
