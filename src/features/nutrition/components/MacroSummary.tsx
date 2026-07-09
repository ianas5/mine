import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';
import { Card } from '@/core/ui';
import type { DayAdherence, MacroSet, MacroStatus } from '@/domain/nutrition';
import type { NutritionTarget } from '@/domain/models';

interface MacroSummaryProps {
  readonly totals: MacroSet;
  readonly target: NutritionTarget | null;
  readonly remaining: MacroSet | null;
  readonly adherence: DayAdherence | null;
}

const g = (value: number): string => (value % 1 === 0 ? String(value) : value.toFixed(1));

function statusColor(theme: Theme, status: MacroStatus | null | undefined): string {
  switch (status) {
    case 'hit':
      return theme.color.positive;
    case 'over':
      return theme.color.danger;
    case 'near':
    case 'under':
      return theme.color.attention;
    default:
      return theme.color.textSecondary;
  }
}

/** Consumed vs. target with remaining and per-macro adherence (FITNESS_DOMAIN §4.2/§4.3). */
export function MacroSummary(props: MacroSummaryProps): ReactNode {
  const theme = useTheme();
  const { totals, target, remaining, adherence } = props;

  const macro = (
    label: string,
    consumed: number,
    targetValue: number | undefined,
    left: number | undefined,
    status: MacroStatus | null,
  ): ReactNode => (
    <View style={{ flex: 1, alignItems: 'center', gap: theme.space.xs }}>
      <Text
        style={{
          ...theme.type.heading,
          color: theme.color.textPrimary,
          fontVariant: ['tabular-nums'],
        }}
      >
        {g(consumed)}
        {targetValue !== undefined ? (
          <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>
            {' '}
            / {g(targetValue)}
          </Text>
        ) : (
          <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}> g</Text>
        )}
      </Text>
      <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>{label}</Text>
      {left !== undefined ? (
        <Text style={{ ...theme.type.micro, color: statusColor(theme, status) }}>
          {left >= 0 ? `${g(left)} left` : `${g(-left)} over`}
        </Text>
      ) : null}
    </View>
  );

  const kcalLeft = remaining?.kcal;

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
          {target ? (
            <Text style={{ ...theme.type.body, color: theme.color.textTertiary }}>
              {' '}
              / {target.kcal.toLocaleString()}
            </Text>
          ) : null}
        </Text>
        <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>CALORIES</Text>
        {kcalLeft !== undefined ? (
          <Text style={{ ...theme.type.caption, color: statusColor(theme, adherence?.calories) }}>
            {kcalLeft >= 0
              ? `${kcalLeft.toLocaleString()} left`
              : `${(-kcalLeft).toLocaleString()} over`}
          </Text>
        ) : (
          <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>
            No target set for this day
          </Text>
        )}
      </View>

      <View style={{ flexDirection: 'row' }}>
        {macro(
          'PROTEIN',
          totals.proteinG,
          target?.proteinG,
          remaining?.proteinG,
          adherence?.protein ?? null,
        )}
        {macro('CARBS', totals.carbG, target?.carbG, remaining?.carbG, adherence?.carbs ?? null)}
        {macro('FAT', totals.fatG, target?.fatG, remaining?.fatG, adherence?.fat ?? null)}
      </View>
    </Card>
  );
}
