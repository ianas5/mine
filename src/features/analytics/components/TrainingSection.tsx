import { useRouter, type Href } from 'expo-router';
import { ChevronRight } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';
import {
  Card,
  ChartFrame,
  EmptyState,
  ProgressBar,
  Skeleton,
  Sparkline,
  StatTile,
  type StatTone,
} from '@/core/ui';
import { formatKg } from '@/core/utils';
import {
  PUSH_PULL_BAND,
  UPPER_LOWER_BAND,
  type Balance,
  type KeyExerciseStrength,
  type RangeKey,
  type Trend,
  type TrendClassification,
  type WorkoutAnalytics,
} from '@/domain/analytics';

import { TrendChart } from './TrendChart';
import { useTrainingAnalytics } from '../hooks/useTrainingAnalytics';

const TONE: Record<TrendClassification, StatTone> = {
  improving: 'positive',
  declining: 'negative',
  stable: 'neutral',
  neutral: 'neutral',
};
const kg1 = (n: number): string => formatKg(Math.round(n * 10) / 10);

export function TrainingSection(props: { readonly range: RangeKey }): ReactNode {
  const theme = useTheme();
  const view = useTrainingAnalytics(props.range);

  return (
    <View style={{ gap: theme.space.lg }}>
      <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>TRAINING</Text>
      {view === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={96} />
          <Skeleton height={160} />
        </View>
      ) : view.workout.totalWorkouts === 0 ? (
        <EmptyState title="No training logged in this range yet." />
      ) : (
        <Loaded workout={view.workout} />
      )}
    </View>
  );
}

function Loaded(props: { readonly workout: WorkoutAnalytics }): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const w = props.workout;
  const { progress, streak } = w.consistency;

  return (
    <View style={{ gap: theme.space.lg }}>
      {/* Consistency — am I training as often as I intend? */}
      <Card style={{ gap: theme.space.sm }}>
        <StatTile
          label="Consistency"
          value={streak > 0 ? `${streak}` : `${progress.completed}`}
          unit={streak > 0 ? 'wk streak' : `of ${progress.planned} this week`}
          context={`${progress.completed} of ${progress.planned} sessions this week · ${round1(w.frequencyPerWeek)}/wk in range`}
          tone={progress.completed >= progress.planned ? 'positive' : 'neutral'}
        />
        <ProgressBar
          value={progress.planned > 0 ? progress.completed / progress.planned : 0}
          tone="attention"
          accessibilityLabel="This week's sessions"
        />
      </Card>

      {/* Am I getting stronger? — the main lifts' e1RM trend */}
      <StrengthSummary keyExercises={w.keyExercises} />

      {/* Am I balanced? */}
      <View style={{ flexDirection: 'row', gap: theme.space.md }}>
        <View style={{ flex: 1 }}>
          <Card>
            <BalanceTile
              label="Push : Pull"
              balance={w.pushPull}
              heavyHigh="Push-heavy"
              heavyLow="Pull-heavy"
              band={PUSH_PULL_BAND}
            />
          </Card>
        </View>
        <View style={{ flex: 1 }}>
          <Card>
            <BalanceTile
              label="Upper : Lower"
              balance={w.upperLower}
              heavyHigh="Upper-heavy"
              heavyLow="Lower-heavy"
              band={UPPER_LOWER_BAND}
            />
          </Card>
        </View>
      </View>

      {/* Which muscles need attention? */}
      {w.mostTrained && w.leastTrained ? (
        <Card style={{ flexDirection: 'row', gap: theme.space.md }}>
          <View style={{ flex: 1 }}>
            <StatTile
              label="Most trained (30d)"
              value={capitalize(w.mostTrained.group)}
              context={`${w.mostTrained.workingSets} working sets`}
              tone="neutral"
            />
          </View>
          <View style={{ flex: 1 }}>
            <StatTile
              label="Least trained (30d)"
              value={capitalize(w.leastTrained.group)}
              context={
                w.leastTrained.workingSets === 0
                  ? 'not trained — worth a look'
                  : `${w.leastTrained.workingSets} working sets`
              }
              tone={w.leastTrained.workingSets === 0 ? 'attention' : 'neutral'}
            />
          </View>
        </Card>
      ) : null}

      {/* Volume as context (quantity, shown neutrally) */}
      <Card>
        <ChartFrame title="Weekly volume" interpretation={volumeInterpretation(w)}>
          {w.volumeSeries.length >= 2 ? (
            <TrendChart series={w.volumeSeries} />
          ) : (
            <Text
              style={{
                ...theme.type.caption,
                color: theme.color.textSecondary,
                paddingVertical: theme.space.lg,
                textAlign: 'center',
              }}
            >
              A few more weeks of training will chart your volume trend.
            </Text>
          )}
        </ChartFrame>
      </Card>

      <Pressable
        onPress={() => router.push('/analytics/muscles' as Href)}
        accessibilityRole="button"
        accessibilityLabel="Muscle report"
        style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
      >
        <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
          <View style={{ flex: 1 }}>
            <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
              Muscle report
            </Text>
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              A per-muscle coaching review
            </Text>
          </View>
          <ChevronRight color={theme.color.textTertiary} size={20} />
        </Card>
      </Pressable>
    </View>
  );
}

function StrengthSummary(props: {
  readonly keyExercises: readonly KeyExerciseStrength[];
}): ReactNode {
  const theme = useTheme();
  const improving = props.keyExercises.filter(
    (k) => k.trend.status === 'ok' && k.trend.value.classification === 'improving',
  ).length;
  const trended = props.keyExercises.filter((k) => k.trend.status === 'ok').length;

  return (
    <Card style={{ gap: theme.space.md }}>
      <StatTile
        label="Getting stronger?"
        value={`${improving}`}
        unit={`of ${trended} main lifts`}
        context={
          trended === 0
            ? 'Log a few more sessions per lift to read a trend'
            : `${improving} of your ${trended} tracked lifts are trending up`
        }
        tone={improving > 0 ? 'positive' : 'neutral'}
      />
      <View style={{ gap: theme.space.sm }}>
        {props.keyExercises.slice(0, 4).map((k) => (
          <View
            key={k.exerciseId}
            style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.sm }}
          >
            <Text style={{ ...theme.type.caption, color: theme.color.textPrimary, flex: 1 }}>
              {k.name}
            </Text>
            {k.trend.status === 'ok' ? (
              <>
                <Sparkline
                  values={k.trend.value.deltaOverWindow >= 0 ? [0, 1] : [1, 0]}
                  color={toneColor(theme, TONE[k.trend.value.classification])}
                  width={40}
                />
                <Text
                  style={{
                    ...theme.type.caption,
                    color: toneColor(theme, TONE[k.trend.value.classification]),
                  }}
                >
                  {k.trend.value.classification === 'stable'
                    ? 'holding'
                    : `${k.trend.value.slopePerWeek > 0 ? '+' : ''}${kg1(k.trend.value.slopePerWeek)} kg/wk`}
                </Text>
              </>
            ) : (
              <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>
                needs more sessions
              </Text>
            )}
          </View>
        ))}
      </View>
    </Card>
  );
}

function BalanceTile(props: {
  readonly label: string;
  readonly balance: Balance;
  readonly heavyHigh: string;
  readonly heavyLow: string;
  readonly band: { readonly low: number; readonly high: number };
}): ReactNode {
  const { balance, band } = props;
  if (balance.ratio === null) {
    return (
      <StatTile label={props.label} value="—" context="not enough data to compare" tone="neutral" />
    );
  }
  const verdict = balance.flagged
    ? balance.ratio > band.high
      ? props.heavyHigh
      : props.heavyLow
    : 'Balanced';
  return (
    <StatTile
      label={props.label}
      value={`${(Math.round(balance.ratio * 10) / 10).toFixed(1)} : 1`}
      context={verdict}
      tone={balance.flagged ? 'attention' : 'positive'}
    />
  );
}

function volumeInterpretation(w: WorkoutAnalytics): string {
  if (w.volumeTrend.status !== 'ok') return w.volumeTrend.needed;
  const t: Trend = w.volumeTrend.value;
  if (t.direction === 'stable') return 'Weekly volume is holding steady';
  return `Weekly volume is ${t.direction === 'increasing' ? 'rising' : 'easing'} (${t.deltaOverWindow > 0 ? '+' : ''}${Math.round(t.deltaOverWindow)} kg over the window)`;
}

function toneColor(theme: Theme, tone: StatTone): string {
  if (tone === 'positive') return theme.color.positive;
  if (tone === 'negative') return theme.color.danger;
  if (tone === 'attention') return theme.color.attention;
  return theme.color.textSecondary;
}

const round1 = (n: number): string => (Math.round(n * 10) / 10).toString();
const capitalize = (s: string): string => s.charAt(0).toUpperCase() + s.slice(1);
