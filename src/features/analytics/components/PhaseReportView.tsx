import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card, SettleIn, StatTile, type StatTone } from '@/core/ui';
import { formatKg } from '@/core/utils';
import {
  type KeyExerciseStrength,
  type PhaseIntentVerdict,
  type PhaseReport,
} from '@/domain/analytics';
import { BODY_FIELD_META, type ChangeDirection, type FieldComparison } from '@/domain/body';
import { PHASE_INTENT, PHASE_TYPE_LABELS } from '@/domain/models';

const round1 = (n: number): number => Math.round(n * 10) / 10;
const signed = (n: number): string => `${n > 0 ? '+' : ''}${round1(n)}`;

const ALIGNMENT_TONE: Record<PhaseIntentVerdict['alignment'], StatTone> = {
  aligned: 'positive',
  counter: 'attention',
  unclear: 'neutral',
};

const DIRECTION_TONE: Record<ChangeDirection, StatTone> = {
  improving: 'positive',
  declining: 'attention',
  stable: 'neutral',
  neutral: 'neutral',
  incomparable: 'neutral',
};

function dayLabel(span: number, weeks: number): string {
  return `${span} days · ${round1(weeks)} weeks`;
}

/** Renders a full Phase Report (ANALYTICS §5.4) — a coach's read of the block, not a
 * spreadsheet: the intent verdict leads, then body change, training quality, nutrition. */
export function PhaseReportView(props: {
  readonly report: PhaseReport;
  /** Set when arriving straight from completing the phase — the report *is* the reward, so
   * its identity header lands with a quiet settle (delight registry #3). */
  readonly celebrate?: boolean;
}): ReactNode {
  const theme = useTheme();
  const { report } = props;
  const { phase, training } = report;

  const header = (
    <View style={{ gap: theme.space.xs }}>
      <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
        {PHASE_TYPE_LABELS[phase.type].toUpperCase()}
        {report.ongoing ? ' · ONGOING' : ''}
      </Text>
      <Text style={{ ...theme.type.title, color: theme.color.textPrimary }}>{phase.name}</Text>
      <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
        {phase.startDate} → {report.ongoing ? 'today' : phase.endDate} ·{' '}
        {dayLabel(report.spanDays, training.weeks)}
      </Text>
      <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>
        Goal: {PHASE_INTENT[phase.type].summary}
      </Text>
    </View>
  );

  return (
    <View style={{ gap: theme.space.xl }}>
      {props.celebrate ? <SettleIn>{header}</SettleIn> : header}

      <IntentBanner verdict={report.nutrition.intent} />

      <Section title="BODY CHANGE">
        <BodyDeltas report={report} />
      </Section>

      <Section title="TRAINING">
        <View style={{ gap: theme.space.lg }}>
          <Card style={{ flexDirection: 'row', gap: theme.space.lg }}>
            <View style={{ flex: 1 }}>
              <StatTile
                label="Workouts"
                value={`${training.workouts}`}
                context={`${round1(training.workoutsPerWeek)} per week`}
                tone="neutral"
              />
            </View>
            <View style={{ flex: 1 }}>
              <StatTile
                label="Consistency"
                value={
                  training.avgWeeklyConsistencyPct.status === 'ok'
                    ? `${training.avgWeeklyConsistencyPct.value}%`
                    : '—'
                }
                context="weekly, vs target"
                tone="neutral"
              />
            </View>
          </Card>

          <Card style={{ gap: theme.space.sm }}>
            <StatTile
              label="Total volume"
              value={formatKg(Math.round(training.totalVolumeKg))}
              unit="kg"
              context={`${formatKg(report.rates.volumePerWeekKg)} kg per week`}
              tone="neutral"
            />
            {training.prCount > 0 ? (
              <Text style={{ ...theme.type.caption, color: theme.color.positive }}>
                {training.prCount} personal record{training.prCount === 1 ? '' : 's'} across{' '}
                {training.prs.length} lift{training.prs.length === 1 ? '' : 's'} this block.
              </Text>
            ) : (
              <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
                No new personal records in this block.
              </Text>
            )}
          </Card>

          {training.keyExercises.length > 0 ? (
            <Card style={{ gap: theme.space.sm }}>
              <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
                Getting stronger?
              </Text>
              {training.keyExercises.map((k) => (
                <KeyExerciseRow key={k.exerciseId} k={k} />
              ))}
            </Card>
          ) : null}
        </View>
      </Section>

      <Section title="NUTRITION">
        <NutritionSummary report={report} />
      </Section>
    </View>
  );
}

function IntentBanner(props: { readonly verdict: PhaseIntentVerdict }): ReactNode {
  const theme = useTheme();
  const tone = ALIGNMENT_TONE[props.verdict.alignment];
  const color =
    tone === 'positive'
      ? theme.color.positive
      : tone === 'attention'
        ? theme.color.attention
        : theme.color.textSecondary;
  return (
    <Card style={{ borderLeftWidth: 3, borderLeftColor: color, gap: theme.space.xs }}>
      <Text style={{ ...theme.type.micro, color }}>
        {props.verdict.alignment === 'aligned'
          ? 'ON TRACK'
          : props.verdict.alignment === 'counter'
            ? 'WORTH A LOOK'
            : 'NOT ENOUGH DATA'}
      </Text>
      <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
        {props.verdict.message}
      </Text>
    </Card>
  );
}

function BodyDeltas(props: { readonly report: PhaseReport }): ReactNode {
  const theme = useTheme();
  const body = props.report.bodyDeltas;
  if (body.status !== 'ok') {
    return (
      <Card>
        <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>{body.needed}</Text>
      </Card>
    );
  }
  const moved = body.value.fields.filter(
    (f) => f.deltaAbs !== null && (f.field === 'weightKg' || f.direction !== 'stable'),
  );
  if (moved.length === 0) {
    return (
      <Card>
        <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
          Body measurements held steady across this block.
        </Text>
      </Card>
    );
  }
  return (
    <Card style={{ gap: theme.space.sm }}>
      {moved.map((f) => (
        <DeltaRow key={f.field} f={f} />
      ))}
    </Card>
  );
}

function DeltaRow(props: { readonly f: FieldComparison }): ReactNode {
  const theme = useTheme();
  const { f } = props;
  const meta = BODY_FIELD_META[f.field];
  const tone = DIRECTION_TONE[f.direction];
  const color =
    tone === 'positive'
      ? theme.color.positive
      : tone === 'attention'
        ? theme.color.attention
        : theme.color.textPrimary;
  return (
    <View style={{ flexDirection: 'row', justifyContent: 'space-between' }}>
      <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>{meta.label}</Text>
      <Text style={{ ...theme.type.bodyStrong, color, fontVariant: ['tabular-nums'] }}>
        {signed(f.deltaAbs ?? 0)} {meta.unit}
      </Text>
    </View>
  );
}

function KeyExerciseRow(props: { readonly k: KeyExerciseStrength }): ReactNode {
  const theme = useTheme();
  const { trend } = props.k;
  let text: string;
  let color = theme.color.textSecondary;
  if (trend.status !== 'ok') {
    text = 'needs more sessions';
  } else if (trend.value.classification === 'improving') {
    text = `up ${round1(Math.abs(trend.value.slopePerWeek))} kg/wk`;
    color = theme.color.positive;
  } else if (trend.value.classification === 'declining') {
    text = `down ${round1(Math.abs(trend.value.slopePerWeek))} kg/wk`;
    color = theme.color.attention;
  } else {
    text = 'holding steady';
  }
  return (
    <View style={{ flexDirection: 'row', justifyContent: 'space-between' }}>
      <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>{props.k.name}</Text>
      <Text style={{ ...theme.type.caption, color }}>{text}</Text>
    </View>
  );
}

function NutritionSummary(props: { readonly report: PhaseReport }): ReactNode {
  const theme = useTheme();
  const n = props.report.nutrition.nutrition;
  if (n.loggedDays === 0) {
    return (
      <Card>
        <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
          No meals were logged during this block.
        </Text>
      </Card>
    );
  }
  const tone = (pct: number): StatTone =>
    pct >= 80 ? 'positive' : pct >= 50 ? 'neutral' : 'attention';
  const cal = n.calorieAdherence;
  const pro = n.proteinAdherence;
  return (
    <Card style={{ gap: theme.space.lg }}>
      <View style={{ flexDirection: 'row', gap: theme.space.lg }}>
        <View style={{ flex: 1 }}>
          <StatTile
            label="Calories"
            value={cal.status === 'ok' ? `${cal.value.pct}%` : '—'}
            context="logged days on target"
            tone={cal.status === 'ok' ? tone(cal.value.pct) : 'neutral'}
          />
        </View>
        <View style={{ flex: 1 }}>
          <StatTile
            label="Protein"
            value={pro.status === 'ok' ? `${pro.value.pct}%` : '—'}
            context="logged days on target"
            tone={pro.status === 'ok' ? tone(pro.value.pct) : 'neutral'}
          />
        </View>
      </View>
      <StatTile
        label="Logging completeness"
        value={`${n.loggedDays}`}
        unit={`of ${n.daysInRange} days`}
        context={
          n.avg
            ? `Averaging ${Math.round(n.avg.kcal)} kcal · ${Math.round(n.avg.proteinG)} g protein on logged days`
            : 'Keep logging for reliable numbers'
        }
        tone="neutral"
      />
    </Card>
  );
}

function Section(props: { readonly title: string; readonly children: ReactNode }): ReactNode {
  const theme = useTheme();
  return (
    <View style={{ gap: theme.space.md }}>
      <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>{props.title}</Text>
      {props.children}
    </View>
  );
}
