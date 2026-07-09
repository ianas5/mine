import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';

/**
 * The optimistic in-session PR marker (UI_UX §4.1, delight registry). A quiet
 * accent pill — the "whole party" is this badge + one success haptic + the finish
 * toast (P18: satisfying, never manipulative). Purely a hint; the authoritative
 * record is always recomputed from history.
 */
export function PrBadge(): ReactNode {
  const theme = useTheme();
  return (
    <View
      style={{
        paddingHorizontal: theme.space.xs,
        paddingVertical: theme.space.xs,
        borderRadius: theme.radius.full,
        backgroundColor: theme.color.accent,
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      <Text
        style={{
          ...theme.type.micro,
          color: theme.color.bg,
          lineHeight: theme.type.micro.fontSize,
        }}
      >
        PR
      </Text>
    </View>
  );
}
