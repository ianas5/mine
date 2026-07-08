import { Ruler } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { useTheme } from '@/core/theme';

export function MeasurementsScreen(): ReactNode {
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
      <Ruler color={theme.color.accent} size={32} strokeWidth={1.75} />
      <Text style={{ ...theme.type.title, color: theme.color.textPrimary }}>Measurements</Text>
      <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
        Body tracking arrives in Phases 11–13
      </Text>
    </View>
  );
}
