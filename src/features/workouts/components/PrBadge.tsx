import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { SettleIn } from '@/core/ui';

/**
 * The optimistic in-session PR marker (UI_UX §4.1, delight registry #1). A quiet accent
 * pill that materializes with a soft scale-settle — the "whole party" is this badge + one
 * success haptic + the finish toast (P18: satisfying, never manipulative). Purely a hint;
 * the authoritative record is always recomputed from history. Settle is skipped under
 * Reduce Motion (the badge still appears — feedback without motion).
 */
export function PrBadge(): ReactNode {
  const theme = useTheme();
  return (
    <SettleIn>
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
    </SettleIn>
  );
}
