import type { ReactNode } from 'react';
import { useState } from 'react';
import { Text, View } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';
import {
  Card,
  ChartFrame,
  EmptyState,
  Screen,
  SegmentedControl,
  Skeleton,
  StatTile,
  type StatTone,
} from '@/core/ui';
import { formatKg } from '@/core/utils';
import {
  RANGE_KEYS,
  RANGE_LABELS,
  type MetricResult,
  type RangeKey,
  type Trend,
  type TrendClassification,
} from '@/domain/analytics';

import { InsightList } from '../components/InsightList';
import { NutritionSection } from '../components/NutritionSection';
import { TrainingSection } from '../components/TrainingSection';
import { TrendChart } from '../components/TrendChart';
import { useBodyAnalytics, type BodyAnalyticsView } from '../hooks/useBodyAnalytics';

const RANGE_OPTIONS = RANGE_KEYS.map((k) => RANGE_LABELS[k]);

const CLASSIFICATION_TONE: Record<TrendClassification, StatTone> = {
  improving: 'positive',
  declining: 'negative',
  stable: 'neutral',
  neutral: 'neutral',
};

const kg1 = (n: number): string => formatKg(Math.round(n * 10) / 10);
const cm1 = (n: number): string => (Math.round(n * 10) / 10).toString();

function rateLabel(trend: Trend, unit: string): string {
  if (trend.direction === 'stable') return 'holding steady';
  const verb = trend.direction === 'increasing' ? 'up' : 'down';
  return `${verb} ${cm1(Math.abs(trend.slopePerWeek))} ${unit}/week`;
}

/**
 * The Analytics home (ANALYTICS §5.3). M2's Body section is live; Training & Nutrition
 * sections arrive in Phases 17–18, shown as honest placeholders — never faked (P8).
 */
export function AnalyticsScreen(): ReactNode {
  const theme = useTheme();
  const [rangeIndex, setRangeIndex] = useState(1); // default 30d
  const range: RangeKey = RANGE_KEYS[rangeIndex] ?? '30d';
  const view = useBodyAnalytics(range);

  return (
    <Screen scroll>
      <Text
        style={{
          ...theme.type.title,
          color: theme.color.textPrimary,
          marginTop: theme.space.sm,
          marginBottom: theme.space.lg,
        }}
      >
        Analytics
      </Text>

      <View style={{ marginBottom: theme.space.lg }}>
        <InsightList />
      </View>

      <View style={{ marginBottom: theme.space.lg }}>
        <SegmentedControl
          options={RANGE_OPTIONS}
          selectedIndex={rangeIndex}
          onChange={setRangeIndex}
          accessibilityLabel="Time range"
        />
      </View>

      {view === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={220} />
          <Skeleton height={120} />
        </View>
      ) : (
        <BodySection view={view} />
      )}

      <View style={{ marginTop: theme.space.xl }}>
        <TrainingSection range={range} />
      </View>

      <View style={{ marginTop: theme.space.xl }}>
        <NutritionSection range={range} />
      </View>
    </Screen>
  );
}

function BodySection(props: { readonly view: BodyAnalyticsView }): ReactNode {
  const theme = useTheme();
  const { analytics, weightSeries } = props.view;
  const { weight, weightTrend, distanceToTarget } = analytics;

  return (
    <View style={{ gap: theme.space.xl }}>
      <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>BODY</Text>

      {weight.latestKg === null ? (
        <EmptyState title="Log a few weigh-ins to see your weight trend." />
      ) : (
        <Card style={{ gap: theme.space.lg }}>
          <ChartFrame
            title="Weight"
            interpretation={weightInterpretation(analytics)}
            interpretationTone={
              weightTrend.status === 'ok'
                ? CLASSIFICATION_TONE[weightTrend.value.classification]
                : 'neutral'
            }
          >
            {weightSeries.length >= 2 ? (
              <TrendChart
                series={weightSeries}
                targetValue={
                  distanceToTarget.status === 'ok' ? distanceToTarget.value.targetKg : null
                }
              />
            ) : (
              <NeedMore theme={theme} text="Not enough weigh-ins in this range to chart yet." />
            )}
          </ChartFrame>

          <View style={{ flexDirection: 'row', gap: theme.space.lg }}>
            <View style={{ flex: 1 }}>
              <StatTile
                label="Trend weight"
                value={kg1(weight.trendKg ?? weight.latestKg)}
                unit="kg"
                context={
                  weightTrend.status === 'ok'
                    ? rateLabel(weightTrend.value, 'kg')
                    : `latest ${kg1(weight.latestKg)} kg`
                }
                tone={
                  weightTrend.status === 'ok'
                    ? CLASSIFICATION_TONE[weightTrend.value.classification]
                    : 'neutral'
                }
              />
            </View>
            <View style={{ flex: 1 }}>
              <DistanceTile distance={distanceToTarget} />
            </View>
          </View>
        </Card>
      )}

      <SiteTrends view={props.view} />
    </View>
  );
}

function DistanceTile(props: {
  readonly distance: BodyAnalyticsView['analytics']['distanceToTarget'];
}): ReactNode {
  const theme = useTheme();
  const d = props.distance;

  if (d.status !== 'ok') {
    return (
      <View style={{ gap: theme.space.xs }}>
        <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>TO GOAL</Text>
        <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
          {d.reason === 'no-target-set' ? 'Set a goal weight in Settings' : d.needed}
        </Text>
      </View>
    );
  }

  if (d.value.atGoal) {
    return (
      <StatTile
        label="To goal"
        value="At goal"
        context="within range of your target"
        tone="positive"
      />
    );
  }
  const above = d.value.toGoKg > 0;
  const eta = d.value.etaWeeks;
  return (
    <StatTile
      label="To goal"
      value={kg1(Math.abs(d.value.toGoKg))}
      unit="kg"
      context={`${above ? 'above' : 'below'} goal${eta !== null ? ` · ~${eta} wk at pace` : ''}`}
      tone={eta !== null ? 'positive' : 'neutral'}
    />
  );
}

function SiteTrends(props: { readonly view: BodyAnalyticsView }): ReactNode {
  const waist = props.view.analytics.siteTrends.get('waistCm');
  if (!waist || waist.latest === null) return null;

  return (
    <Card>
      <StatTile
        label="Waist"
        value={cm1(waist.latest)}
        unit="cm"
        context={
          waist.trend.status === 'ok' ? rateLabel(waist.trend.value, 'cm') : waist.trend.needed
        }
        tone={
          waist.trend.status === 'ok'
            ? CLASSIFICATION_TONE[waist.trend.value.classification]
            : 'neutral'
        }
      />
    </Card>
  );
}

function NeedMore(props: { readonly theme: Theme; readonly text: string }): ReactNode {
  return (
    <View
      style={{
        height: 120,
        alignItems: 'center',
        justifyContent: 'center',
        paddingHorizontal: props.theme.space.lg,
      }}
    >
      <Text
        style={{
          ...props.theme.type.caption,
          color: props.theme.color.textSecondary,
          textAlign: 'center',
        }}
      >
        {props.text}
      </Text>
    </View>
  );
}

function weightInterpretation(analytics: BodyAnalyticsView['analytics']): string {
  const wt: MetricResult<Trend> = analytics.weightTrend;
  if (wt.status !== 'ok') return wt.needed;
  const trendKg = analytics.weight.trendKg ?? analytics.weight.latestKg;
  const head = trendKg !== null ? `Trend weight ${kg1(trendKg)} kg` : 'Weight';
  return `${head} — ${rateLabel(wt.value, 'kg')}`;
}
