import { useRouter, type Href } from 'expo-router';
import {
  Camera,
  ChevronRight,
  Dumbbell,
  Flame,
  Ruler,
  Scale,
  Settings,
  UtensilsCrossed,
  type LucideIcon,
} from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';
import { Card, IconButton, InsightCard, ProgressBar, Screen, Skeleton } from '@/core/ui';
import { insightEvidenceHref, useInsights } from '@/data/analytics/useInsights';

import { Ring } from '../components/Ring';
import { useDashboard, type DashboardData } from '../hooks/useDashboard';

/** A minimal view of the live session, supplied by the composition root (app route) so
 * the dashboard never imports the workouts store across the feature boundary. */
export interface ActiveSessionSummary {
  readonly name: string;
  readonly startedAt: number | null;
  readonly exerciseCount: number;
  readonly completedSetCount: number;
}

interface DashboardScreenProps {
  readonly activeSession?: ActiveSessionSummary | null;
}

const MORNING_END_HOUR = 11; // "morning" = before ~11:00 (UI_UX §7.6)

export function DashboardScreen(props: DashboardScreenProps): ReactNode {
  const theme = useTheme();
  const data = useDashboard();
  const active = props.activeSession ?? null;
  const isMorning = new Date().getHours() < MORNING_END_HOUR;

  return (
    <Screen scroll>
      <Header data={data} isMorning={isMorning} />

      {data === undefined ? (
        <View style={{ gap: theme.space.md, marginTop: theme.space.md }}>
          <Skeleton height={96} />
          <Skeleton height={160} />
        </View>
      ) : (
        <View style={{ gap: theme.space.lg, marginTop: theme.space.md }}>
          {active ? (
            <LiveSessionCard session={active} />
          ) : isMorning ? (
            <>
              <WorkoutCard data={data} />
              <MacrosCard data={data} />
            </>
          ) : (
            <>
              <MacrosCard data={data} />
              <WorkoutCard data={data} />
            </>
          )}

          {/* Insight slot (ANALYTICS §6.5) — top-3, hidden during Focus Mode (§5.1). */}
          {active ? null : <InsightSlot />}

          <StreakLine data={data} />
        </View>
      )}

      <QuickActions focus={active !== null} />
      <View style={{ height: theme.space.xl }} />
    </Screen>
  );
}

function Header(props: {
  readonly data: DashboardData | undefined;
  readonly isMorning: boolean;
}): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const { data, isMorning } = props;

  const greeting = isMorning
    ? 'Good morning'
    : new Date().getHours() < 18
      ? 'Good afternoon'
      : 'Good evening';

  let sub: string;
  if (!data) sub = ' ';
  else if (
    data.greeting.daysSinceLastWorkout !== null &&
    data.greeting.daysSinceLastWorkout >= 14
  ) {
    const weeks = Math.round(data.greeting.daysSinceLastWorkout / 7);
    sub = `Back at it — first session in ${weeks} week${weeks === 1 ? '' : 's'}`;
  } else if (isMorning && !data.greeting.weighedInToday) {
    sub = 'Weigh in to update your trend';
  } else if (data.greeting.trendWeightKg !== null) {
    sub = `Trend weight ${roundKg(data.greeting.trendWeightKg)} kg`;
  } else {
    sub = 'Log a weigh-in to start your trend';
  }

  return (
    <View
      style={{
        flexDirection: 'row',
        alignItems: 'flex-start',
        justifyContent: 'space-between',
        marginTop: theme.space.sm,
      }}
    >
      <View style={{ flex: 1, gap: theme.space.xs }}>
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary }}>{greeting}</Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>{sub}</Text>
      </View>
      <IconButton
        icon={<Settings color={theme.color.textSecondary} size={24} strokeWidth={1.75} />}
        onPress={() => router.push('/settings')}
        accessibilityLabel="Settings"
      />
    </View>
  );
}

function LiveSessionCard(props: { readonly session: ActiveSessionSummary }): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const { session } = props;
  return (
    <Pressable
      onPress={() => router.push('/active-workout')}
      accessibilityRole="button"
      accessibilityLabel="Return to workout"
      style={({ pressed }) => ({ opacity: pressed ? 0.8 : 1 })}
    >
      <Card
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.md,
          borderColor: theme.color.accent,
          borderWidth: 1,
        }}
      >
        <Dumbbell color={theme.color.accent} size={22} strokeWidth={1.75} />
        <View style={{ flex: 1 }}>
          <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
            {session.name} in progress
          </Text>
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            {session.exerciseCount} exercise{session.exerciseCount === 1 ? '' : 's'} ·{' '}
            {session.completedSetCount} set{session.completedSetCount === 1 ? '' : 's'} logged
          </Text>
        </View>
        <Text style={{ ...theme.type.caption, color: theme.color.accent }}>Return</Text>
        <ChevronRight color={theme.color.accent} size={20} />
      </Card>
    </Pressable>
  );
}

function WorkoutCard(props: { readonly data: DashboardData }): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const { workout } = props.data;

  const title =
    workout.state === 'done'
      ? 'Workout done today'
      : workout.state === 'planned'
        ? (workout.suggestionLabel ?? 'Ready to train')
        : 'Rest day';
  const sub =
    workout.state === 'done'
      ? 'Nice — logged and counted'
      : workout.state === 'planned'
        ? "Today's suggested session"
        : 'Nothing scheduled — start anything when you like';

  return (
    <Pressable
      onPress={() => router.push('/workouts' as Href)}
      accessibilityRole="button"
      accessibilityLabel={`Today's workout: ${title}`}
      style={({ pressed }) => ({ opacity: pressed ? 0.8 : 1 })}
    >
      <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
        <Dumbbell
          color={workout.state === 'done' ? theme.color.positive : theme.color.accent}
          size={22}
          strokeWidth={1.75}
        />
        <View style={{ flex: 1 }}>
          <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>{title}</Text>
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>{sub}</Text>
        </View>
        <ChevronRight color={theme.color.textTertiary} size={20} />
      </Card>
    </Pressable>
  );
}

function MacrosCard(props: { readonly data: DashboardData }): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const { target, totals, remaining } = props.data.nutrition;

  if (target === null || remaining === null) {
    return (
      <Pressable
        onPress={() => router.push('/nutrition' as Href)}
        accessibilityRole="button"
        accessibilityLabel="Set nutrition targets"
        style={({ pressed }) => ({ opacity: pressed ? 0.8 : 1 })}
      >
        <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
          <Flame color={theme.color.accent} size={22} strokeWidth={1.75} />
          <View style={{ flex: 1 }}>
            <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
              Set your nutrition targets
            </Text>
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              No target set for today — tap to add one
            </Text>
          </View>
          <ChevronRight color={theme.color.textTertiary} size={20} />
        </Card>
      </Pressable>
    );
  }

  const kcalLeft = remaining.kcal;
  const proteinLeft = remaining.proteinG;

  return (
    <Card style={{ gap: theme.space.lg }}>
      <View style={{ flexDirection: 'row', justifyContent: 'space-around' }}>
        <Ring
          label="calories left"
          value={String(Math.round(kcalLeft))}
          unit="kcal"
          fraction={fraction(totals.kcal, target.kcal)}
          tone={kcalLeft < 0 ? 'attention' : 'neutral'}
        />
        <Ring
          label="protein left"
          value={`${Math.max(0, Math.round(proteinLeft))}`}
          unit="g"
          fraction={fraction(totals.proteinG, target.proteinG)}
          tone={proteinLeft <= 0 ? 'positive' : 'neutral'}
        />
      </View>
      <View style={{ gap: theme.space.sm }}>
        <MacroBar label="Carbs" consumed={totals.carbG} target={target.carbG} theme={theme} />
        <MacroBar label="Fat" consumed={totals.fatG} target={target.fatG} theme={theme} />
      </View>
    </Card>
  );
}

function MacroBar(props: {
  readonly label: string;
  readonly consumed: number;
  readonly target: number;
  readonly theme: Theme;
}): ReactNode {
  const { label, consumed, target, theme } = props;
  return (
    <View style={{ gap: theme.space.xs }}>
      <View style={{ flexDirection: 'row', justifyContent: 'space-between' }}>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>{label}</Text>
        <Text
          style={{
            ...theme.type.caption,
            color: theme.color.textTertiary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {Math.round(consumed)} / {Math.round(target)} g
        </Text>
      </View>
      <ProgressBar value={fraction(consumed, target)} accessibilityLabel={`${label} progress`} />
    </View>
  );
}

function StreakLine(props: { readonly data: DashboardData }): ReactNode {
  const theme = useTheme();
  const { weeks, progress } = props.data.streak;
  return (
    <Card style={{ gap: theme.space.sm }}>
      <View style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.sm }}>
        <Flame
          color={weeks > 0 ? theme.color.attention : theme.color.textTertiary}
          size={20}
          strokeWidth={1.75}
        />
        <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary, flex: 1 }}>
          {weeks > 0 ? `${weeks}-week streak` : 'No active streak — this week counts'}
        </Text>
        <Text
          style={{
            ...theme.type.caption,
            color: theme.color.textSecondary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {progress.completed} of {progress.planned} this week
        </Text>
      </View>
      <ProgressBar
        value={progress.planned > 0 ? progress.completed / progress.planned : 0}
        tone="attention"
        accessibilityLabel="This week's workouts"
      />
    </Card>
  );
}

interface Action {
  readonly key: string;
  readonly label: string;
  readonly icon: LucideIcon;
  readonly href: Href;
}

const ALL_ACTIONS: readonly Action[] = [
  { key: 'start', label: 'Start', icon: Dumbbell, href: '/workouts' as Href },
  { key: 'meal', label: 'Log Meal', icon: UtensilsCrossed, href: '/nutrition?open=meal' as Href },
  { key: 'weight', label: 'Weight', icon: Scale, href: '/measurements?open=weight' as Href },
  { key: 'measure', label: 'Measure', icon: Ruler, href: '/measurements?open=measure' as Href },
  { key: 'photo', label: 'Photo', icon: Camera, href: '/measurements/photos' as Href },
];

// Focus Mode (§5.1): Start → Return to Workout; Log Meal & Add Weight stay; Measure/Photo drop.
const FOCUS_ACTIONS: readonly Action[] = [
  { key: 'return', label: 'Return', icon: Dumbbell, href: '/active-workout' as Href },
  { key: 'meal', label: 'Log Meal', icon: UtensilsCrossed, href: '/nutrition?open=meal' as Href },
  { key: 'weight', label: 'Weight', icon: Scale, href: '/measurements?open=weight' as Href },
];

function QuickActions(props: { readonly focus: boolean }): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const actions = props.focus ? FOCUS_ACTIONS : ALL_ACTIONS;

  return (
    <View style={{ flexDirection: 'row', gap: theme.space.sm, marginTop: theme.space.lg }}>
      {actions.map((action) => {
        const Icon = action.icon;
        return (
          <Pressable
            key={action.key}
            onPress={() => router.push(action.href)}
            accessibilityRole="button"
            accessibilityLabel={action.label}
            style={({ pressed }) => ({
              flex: 1,
              alignItems: 'center',
              gap: theme.space.xs,
              paddingVertical: theme.space.md,
              borderRadius: theme.radius.md,
              backgroundColor: theme.color.surface,
              borderWidth: 1,
              borderColor: theme.color.border,
              opacity: pressed ? 0.7 : 1,
            })}
          >
            <Icon color={theme.color.accent} size={22} strokeWidth={1.75} />
            <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
              {action.label}
            </Text>
          </Pressable>
        );
      })}
    </View>
  );
}

function InsightSlot(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const insights = useInsights();
  if (!insights || insights.dashboard.length === 0) return null;

  return (
    <View style={{ gap: theme.space.sm }}>
      {insights.dashboard.map((insight) => (
        <InsightCard
          key={insight.instanceKey}
          tone={insight.tone}
          title={insight.title}
          body={insight.body}
          onPress={() => router.push(insightEvidenceHref(insight.evidence) as Href)}
          onDismiss={() => insights.dismiss(insight.instanceKey, insight.classification)}
        />
      ))}
    </View>
  );
}

const roundKg = (n: number): string => (Math.round(n * 10) / 10).toString();
const fraction = (consumed: number, target: number): number => (target > 0 ? consumed / target : 0);
