import { ChevronRight } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';

interface ListRowProps {
  readonly title: string;
  readonly subtitle?: string;
  readonly leading?: ReactNode;
  readonly trailingValue?: string;
  readonly chevron?: boolean;
  readonly onPress?: () => void;
}

/** List row — 52pt minimum height, hairline-separated by the parent list. */
export function ListRow(props: ListRowProps): ReactNode {
  const theme = useTheme();
  const content = (
    <View
      style={{
        minHeight: 52,
        flexDirection: 'row',
        alignItems: 'center',
        gap: theme.space.md,
        paddingVertical: theme.space.sm,
      }}
    >
      {props.leading ? <View>{props.leading}</View> : null}
      <View style={{ flex: 1 }}>
        <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>{props.title}</Text>
        {props.subtitle ? (
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            {props.subtitle}
          </Text>
        ) : null}
      </View>
      {props.trailingValue ? (
        <Text
          style={{
            ...theme.type.bodyStrong,
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {props.trailingValue}
        </Text>
      ) : null}
      {props.chevron ? <ChevronRight color={theme.color.textTertiary} size={20} /> : null}
    </View>
  );
  if (!props.onPress) {
    return content;
  }
  return (
    <Pressable
      onPress={props.onPress}
      accessibilityRole="button"
      accessibilityLabel={props.title}
      style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
    >
      {content}
    </Pressable>
  );
}
