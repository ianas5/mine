import { useLocalSearchParams, useRouter } from 'expo-router';
import { ArrowLeft } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';
import {
  Card,
  EmptyState,
  IconButton,
  Screen,
  Skeleton,
  Sparkline,
  type StatTone,
} from '@/core/ui';
import { formatKg, formatRelativeDate } from '@/core/utils';
import type { ExerciseReport, ExerciseTrend, TrendClassification } from '@/domain/analytics';

import { useExerciseReport } from '../hooks/useExerciseReport';

const CLASSIFICATION_TONE: Record<TrendClassification, StatTone> = {
  improving: 'positive',
  declining: 'negative',
  stable: 'neutral',
  neutral: 'neutral',
};

function toneColor(theme: Theme, tone: StatTone): string {
  if (tone === 'positive') return theme.color.positive;
  if (tone === 'negative') return theme.color.danger;
  if (tone === 'attention') return theme.color.attention;
  return theme.color.textSecondary;
}

const kg = (value: number | null): string => (value === null ? '—' : `${formatKg(value)} kg`);
/** e1RM displays to the nearest 0.5 kg (FITNESS_DOMAIN §3.5); raw is stored elsewhere. */
const e1rm = (value: number | null): string =>
  value === null ? '—' : `${formatKg(Math.round(value * 2) / 2)} kg`;
const volume = (value: number | null): string =>
  value === null ? '—' : `${Math.round(value).toLocaleString()} kg`;

/** Per-exercise report (ANALYTICS §5.5): bests, totals, averages, last performed. */
export function ExerciseReportScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const params = useLocalSearchParams<{ id: string }>();
  const view = useExerciseReport(params.id ?? '');

  return (
    <Screen scroll>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.sm,
          marginTop: theme.space.sm,
          marginBottom: theme.space.lg,
        }}
      >
        <IconButton
          icon={<ArrowLeft color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
          onPress={() => router.back()}
          accessibilityLabel="Back"
        />
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary, flex: 1 }}>
          {view ? view.name : 'Exercise'}
        </Text>
      </View>

      {view === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={96} />
          <Skeleton height={96} />
        </View>
      ) : view === null ? (
        <EmptyState title="Exercise not found" cta={{ label: 'Go back', onPress: router.back }} />
      ) : view.report.totalWorkingSets === 0 ? (
        <EmptyState title="No working sets logged yet — the report fills in as you train." />
      ) : (
        <LoadedReport report={view.report} trend={view.trend} />
      )}
    </Screen>
  );
}

function LoadedReport(props: {
  readonly report: ExerciseReport;
  readonly trend: ExerciseTrend;
}): ReactNode {
  const theme = useTheme();
  const { report, trend } = props;

  const cell = (label: string, value: string): ReactNode => (
    <View style={{ flex: 1, gap: theme.space.xs }}>
      <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>{label}</Text>
      <Text
        style={{
          ...theme.type.heading,
          color: theme.color.textPrimary,
          fontVariant: ['tabular-nums'],
        }}
      >
        {value}
      </Text>
    </View>
  );

  const row = (a: ReactNode, b: ReactNode): ReactNode => (
    <View style={{ flexDirection: 'row', gap: theme.space.md }}>
      {a}
      {b}
    </View>
  );

  return (
    <View style={{ gap: theme.space.lg }}>
      {report.lastPerformed ? (
        <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
          Last performed {formatRelativeDate(report.lastPerformed.date)}
        </Text>
      ) : null}

      <View style={{ gap: theme.space.sm }}>
        <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
          PERSONAL RECORDS
        </Text>
        <Card style={{ gap: theme.space.lg }}>
          {row(
            cell('Heaviest weight', kg(report.bests.heaviestWeightKg)),
            cell('Best e1RM', e1rm(report.bests.bestE1rmKg)),
          )}
          {row(
            cell('Best set volume', volume(report.bests.bestSetVolumeKg)),
            cell('Best session volume', volume(report.bests.bestSessionVolumeKg)),
          )}
        </Card>
      </View>

      <View style={{ gap: theme.space.sm }}>
        <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>ALL-TIME</Text>
        <Card style={{ gap: theme.space.lg }}>
          {row(
            cell('Sessions', String(report.totalSessions)),
            cell('Working sets', String(report.totalWorkingSets)),
          )}
          {row(
            cell('Total volume', volume(report.totalVolumeKg)),
            cell(
              'Avg reps / set',
              report.avgRepsPerWorkingSet === null ? '—' : report.avgRepsPerWorkingSet.toFixed(1),
            ),
          )}
          {row(cell('Avg load', kg(report.avgEffectiveLoadKg)), <View style={{ flex: 1 }} />)}
        </Card>
      </View>

      <StrengthTrend trend={trend} />
    </View>
  );
}

/** e1RM strength trend (ANALYTICS §5.1/§5.5): sparkline + progression rate, or the
 * honest "needs more data" state — never a fabricated line. */
function StrengthTrend(props: { readonly trend: ExerciseTrend }): ReactNode {
  const theme = useTheme();
  const { series, trend } = props.trend;

  if (trend.status !== 'ok') {
    return (
      <Card style={{ gap: theme.space.xs }}>
        <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
          Strength trend
        </Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
          {trend.needed}.
        </Text>
      </Card>
    );
  }

  const tone = CLASSIFICATION_TONE[trend.value.classification];
  const rate = trend.value.slopePerWeek;
  const sign = rate > 0 ? '+' : '';
  return (
    <Card style={{ gap: theme.space.sm }}>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          justifyContent: 'space-between',
        }}
      >
        <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
          Strength trend
        </Text>
        <Sparkline
          values={series.map((p) => p.value)}
          color={toneColor(theme, tone)}
          accessibilityLabel="e1RM trend sparkline"
        />
      </View>
      <Text style={{ ...theme.type.caption, color: toneColor(theme, tone) }}>
        e1RM {trend.value.classification} · {sign}
        {formatKg(Math.round(rate * 2) / 2)} kg/week over{' '}
        {trend.value.deltaOverWindow >= 0 ? '+' : ''}
        {Math.round(trend.value.deltaOverWindow)} kg total
      </Text>
    </Card>
  );
}
