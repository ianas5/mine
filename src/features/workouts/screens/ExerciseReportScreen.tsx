import { useLocalSearchParams, useRouter } from 'expo-router';
import { ArrowLeft } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card, EmptyState, IconButton, Screen, Skeleton } from '@/core/ui';
import { formatKg, formatRelativeDate } from '@/core/utils';
import type { ExerciseReport } from '@/domain/analytics';

import { useExerciseReport } from '../hooks/useExerciseReport';

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
        <LoadedReport name={view.name} report={view.report} />
      )}
    </Screen>
  );
}

function LoadedReport(props: {
  readonly name: string;
  readonly report: ExerciseReport;
}): ReactNode {
  const theme = useTheme();
  const { report } = props;

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

      <Card style={{ gap: theme.space.xs }}>
        <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
          Strength trend
        </Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
          The e1RM trend chart and progression rate arrive in a later update. Until then, no trend
          line is shown rather than an estimated one.
        </Text>
      </Card>
    </View>
  );
}
