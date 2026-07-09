import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card, EmptyState, ProgressBar, Skeleton, StatTile, type StatTone } from '@/core/ui';
import type { AdherenceStat, MetricResult, NutritionAnalytics, RangeKey } from '@/domain/analytics';

import { useNutritionAnalytics } from '../hooks/useNutritionAnalytics';

const pct = (n: number): string => `${Math.round(n)}%`;

function adherenceTone(stat: MetricResult<AdherenceStat>): StatTone {
  if (stat.status !== 'ok') return 'neutral';
  if (stat.value.pct >= 80) return 'positive';
  if (stat.value.pct >= 50) return 'neutral';
  return 'attention';
}

export function NutritionSection(props: { readonly range: RangeKey }): ReactNode {
  const theme = useTheme();
  const n = useNutritionAnalytics(props.range);

  return (
    <View style={{ gap: theme.space.lg }}>
      <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>NUTRITION</Text>
      {n === undefined ? (
        <Skeleton height={120} />
      ) : n.loggedDays === 0 ? (
        <EmptyState title="Log a few days of meals to see your nutrition adherence." />
      ) : (
        <Loaded n={n} />
      )}
    </View>
  );
}

function Loaded(props: { readonly n: NutritionAnalytics }): ReactNode {
  const theme = useTheme();
  const { n } = props;

  const adhTile = (label: string, stat: MetricResult<AdherenceStat>): ReactNode =>
    stat.status === 'ok' ? (
      <StatTile
        label={label}
        value={pct(stat.value.pct)}
        context={`hit ${stat.value.hitDays} of ${stat.value.loggedDays} logged days`}
        tone={adherenceTone(stat)}
      />
    ) : (
      <StatTile label={label} value="—" context="not enough logged days yet" tone="neutral" />
    );

  return (
    <View style={{ gap: theme.space.lg }}>
      <Card style={{ flexDirection: 'row', gap: theme.space.lg }}>
        <View style={{ flex: 1 }}>{adhTile('Calories', n.calorieAdherence)}</View>
        <View style={{ flex: 1 }}>{adhTile('Protein', n.proteinAdherence)}</View>
      </Card>

      <Card style={{ gap: theme.space.sm }}>
        <StatTile
          label="Logging completeness"
          value={`${n.loggedDays}`}
          unit={`of ${n.daysInRange} days`}
          context={
            n.calorieSkew
              ? `Calorie misses skew ${n.calorieSkew} — worth a look`
              : n.avg
                ? `Averaging ${Math.round(n.avg.kcal)} kcal · ${Math.round(n.avg.proteinG)} g protein on logged days`
                : 'Keep logging for reliable trends'
          }
          tone="neutral"
        />
        <ProgressBar
          value={n.completeness}
          tone="attention"
          accessibilityLabel="Logging completeness"
        />
      </Card>
    </View>
  );
}
