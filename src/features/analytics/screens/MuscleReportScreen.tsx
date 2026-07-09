import { useRouter } from 'expo-router';
import { ArrowLeft, TrendingUp } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';
import { Card, IconButton, Screen, SegmentedControl, Skeleton } from '@/core/ui';
import { formatKg, formatRelativeDate } from '@/core/utils';
import {
  RANGE_KEYS,
  RANGE_LABELS,
  type MuscleGroupReport,
  type RangeKey,
} from '@/domain/analytics';

import { useTrainingAnalytics } from '../hooks/useTrainingAnalytics';

const RANGE_OPTIONS = RANGE_KEYS.map((k) => RANGE_LABELS[k]);
const kg1 = (n: number): string => formatKg(Math.round(n * 10) / 10);
const capitalize = (s: string): string => s.charAt(0).toUpperCase() + s.slice(1);

/** The Muscle Report (ANALYTICS §5.6) — a coach's read of every muscle group over the
 * last months: what's strongest, what's improving, what's being neglected. */
export function MuscleReportScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const [rangeIndex, setRangeIndex] = useState(3); // default 6M ("the last months")
  const range: RangeKey = RANGE_KEYS[rangeIndex] ?? '180d';
  const view = useTrainingAnalytics(range);

  const trained = view?.muscles.filter((m) => !m.untrained) ?? [];
  const untrained = view?.muscles.filter((m) => m.untrained) ?? [];

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
          Muscle report
        </Text>
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
          <Skeleton height={140} />
          <Skeleton height={140} />
        </View>
      ) : (
        <View style={{ gap: theme.space.lg }}>
          {trained
            .slice()
            .sort((a, b) => b.volume30dKg - a.volume30dKg)
            .map((report) => (
              <GroupCard key={report.group} report={report} theme={theme} />
            ))}

          {untrained.length > 0 ? (
            <Card style={{ gap: theme.space.xs }}>
              <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
                Not trained in this range
              </Text>
              <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
                {untrained.map((m) => capitalize(m.group)).join(' · ')}
              </Text>
            </Card>
          ) : null}
        </View>
      )}
    </Screen>
  );
}

function GroupCard(props: {
  readonly report: MuscleGroupReport;
  readonly theme: Theme;
}): ReactNode {
  const { report, theme } = props;

  return (
    <Card style={{ gap: theme.space.md }}>
      <View style={{ flexDirection: 'row', alignItems: 'center' }}>
        <Text style={{ ...theme.type.heading, color: theme.color.textPrimary, flex: 1 }}>
          {capitalize(report.group)}
        </Text>
        {report.lastTrained ? (
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            {formatRelativeDate(report.lastTrained.date)}
          </Text>
        ) : null}
      </View>

      <View style={{ flexDirection: 'row', gap: theme.space.md }}>
        <Fact
          label="Volume 30d"
          value={`${Math.round(report.volume30dKg).toLocaleString()} kg`}
          theme={theme}
        />
        <Fact label="Sets 30d" value={String(report.workingSets30d)} theme={theme} />
        <Fact label="Freq" value={`${round1(report.frequencyPerWeek)}/wk`} theme={theme} />
      </View>

      {report.strongest ? (
        <Line
          theme={theme}
          text={`Strongest: ${report.strongest.name} · ${kg1(report.strongest.value)} kg e1RM`}
        />
      ) : null}

      {report.fastestImproving && report.fastestImproving.value > 0 ? (
        <View style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.xs }}>
          <TrendingUp color={theme.color.positive} size={16} />
          <Text style={{ ...theme.type.caption, color: theme.color.positive }}>
            Improving fastest: {report.fastestImproving.name} (+{kg1(report.fastestImproving.value)}{' '}
            kg/wk e1RM)
          </Text>
        </View>
      ) : (
        <Line theme={theme} text="Not enough sessions per lift to rank improvement yet" muted />
      )}
    </Card>
  );
}

function Fact(props: {
  readonly label: string;
  readonly value: string;
  readonly theme: Theme;
}): ReactNode {
  const { theme } = props;
  return (
    <View style={{ flex: 1, gap: theme.space.xs }}>
      <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>
        {props.label.toUpperCase()}
      </Text>
      <Text
        style={{
          ...theme.type.bodyStrong,
          color: theme.color.textPrimary,
          fontVariant: ['tabular-nums'],
        }}
      >
        {props.value}
      </Text>
    </View>
  );
}

function Line(props: {
  readonly theme: Theme;
  readonly text: string;
  readonly muted?: boolean;
}): ReactNode {
  return (
    <Text
      style={{
        ...props.theme.type.caption,
        color: props.muted ? props.theme.color.textTertiary : props.theme.color.textSecondary,
      }}
    >
      {props.text}
    </Text>
  );
}

const round1 = (n: number): string => (Math.round(n * 10) / 10).toString();
