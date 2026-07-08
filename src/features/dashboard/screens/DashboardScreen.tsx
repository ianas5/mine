import { useRouter } from 'expo-router';
import { LayoutDashboard, Settings } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { IconButton, Screen } from '@/core/ui';

export function DashboardScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  return (
    <Screen>
      <View
        style={{
          flexDirection: 'row',
          justifyContent: 'flex-end',
          marginTop: theme.space.sm,
        }}
      >
        <IconButton
          icon={<Settings color={theme.color.textSecondary} size={24} strokeWidth={1.75} />}
          onPress={() => router.push('/settings')}
          accessibilityLabel="Settings"
        />
      </View>
      <View
        style={{
          flex: 1,
          alignItems: 'center',
          justifyContent: 'center',
          gap: theme.space.md,
        }}
      >
        <LayoutDashboard color={theme.color.accent} size={32} strokeWidth={1.75} />
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary }}>Dashboard</Text>
        <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
          Daily briefing arrives in Phase 16
        </Text>
      </View>
    </Screen>
  );
}
