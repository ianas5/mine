import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card } from '@/core/ui';
import type { MacroSet } from '@/domain/nutrition';

/** The day's consumed macros (targets/remaining arrive in Phase 10). */
export function MacroSummary(props: { readonly totals: MacroSet }): ReactNode {
  const theme = useTheme();
  const { totals } = props;

  const macro = (label: string, grams: number): ReactNode => (
    <View style={{ flex: 1, alignItems: 'center', gap: theme.space.xs }}>
      <Text
        style={{
          ...theme.type.heading,
          color: theme.color.textPrimary,
          fontVariant: ['tabular-nums'],
        }}
      >
        {grams % 1 === 0 ? grams : grams.toFixed(1)}
        <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}> g</Text>
      </Text>
      <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>{label}</Text>
    </View>
  );

  return (
    <Card style={{ gap: theme.space.lg }}>
      <View style={{ alignItems: 'center', gap: theme.space.xs }}>
        <Text
          style={{
            ...theme.type.display,
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {totals.kcal.toLocaleString()}
        </Text>
        <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>CALORIES</Text>
      </View>
      <View style={{ flexDirection: 'row' }}>
        {macro('PROTEIN', totals.proteinG)}
        {macro('CARBS', totals.carbG)}
        {macro('FAT', totals.fatG)}
      </View>
    </Card>
  );
}
