import { Timer } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';

import { formatCountdown, useRestCountdown } from '../hooks/useRestCountdown';
import { REST_EXTEND_SEC, useRestActions } from '../stores/useRestTimerStore';

/**
 * Slim, non-blocking rest countdown with skip/extend (UI_UX §4.2, P17). Rendered
 * only while a rest is running; never an overlay that blocks logging the next set.
 */
export function RestTimerBar(): ReactNode {
  const theme = useTheme();
  const rest = useRestActions();
  const { running, remainingMs, durationSec } = useRestCountdown();

  if (!running) return null;

  const progress = durationSec > 0 ? 1 - remainingMs / (durationSec * 1000) : 1;

  return (
    <View
      style={{
        borderRadius: theme.radius.md,
        backgroundColor: theme.color.accentSoft,
        overflow: 'hidden',
      }}
    >
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.sm,
          paddingHorizontal: theme.space.md,
          paddingVertical: theme.space.sm,
        }}
      >
        <Timer color={theme.color.accent} size={18} strokeWidth={2} />
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>Rest</Text>
        <Text
          style={{
            ...theme.type.heading,
            color: theme.color.accent,
            fontVariant: ['tabular-nums'],
            flex: 1,
          }}
        >
          {formatCountdown(remainingMs)}
        </Text>

        <Pressable
          onPress={() => rest.extend(REST_EXTEND_SEC, Date.now())}
          accessibilityRole="button"
          accessibilityLabel={`Add ${REST_EXTEND_SEC} seconds`}
          hitSlop={theme.space.sm}
          style={({ pressed }) => ({
            paddingHorizontal: theme.space.md,
            paddingVertical: theme.space.xs,
            borderRadius: theme.radius.full,
            borderWidth: 1,
            borderColor: theme.color.accent,
            opacity: pressed ? 0.6 : 1,
          })}
        >
          <Text style={{ ...theme.type.caption, color: theme.color.accent }}>
            +{REST_EXTEND_SEC}s
          </Text>
        </Pressable>

        <Pressable
          onPress={() => rest.skip()}
          accessibilityRole="button"
          accessibilityLabel="Skip rest"
          hitSlop={theme.space.sm}
          style={({ pressed }) => ({
            paddingHorizontal: theme.space.md,
            paddingVertical: theme.space.xs,
            opacity: pressed ? 0.6 : 1,
          })}
        >
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>Skip</Text>
        </Pressable>
      </View>

      <View style={{ height: 3, backgroundColor: theme.color.border }}>
        <View
          style={{
            height: 3,
            width: `${Math.round(Math.min(1, Math.max(0, progress)) * 100)}%`,
            backgroundColor: theme.color.accent,
          }}
        />
      </View>
    </View>
  );
}
