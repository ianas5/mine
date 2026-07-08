import { LayoutDashboard } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { useTheme } from '@/core/theme';

export function DashboardScreen(): ReactNode {
  const theme = useTheme();
  const insets = useSafeAreaInsets();
  return (
    <View
      style={{
        flex: 1,
        backgroundColor: theme.color.bg,
        alignItems: 'center',
        justifyContent: 'center',
        paddingTop: insets.top,
        gap: theme.space.md,
      }}
    >
      <LayoutDashboard color={theme.color.accent} size={32} strokeWidth={1.75} />
      <Text style={{ ...theme.type.title, color: theme.color.textPrimary }}>Dashboard</Text>
      <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
        Daily briefing arrives in Phase 16
      </Text>
    </View>
  );
}
