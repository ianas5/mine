import { Settings, TrendingDown, TrendingUp } from 'lucide-react-native';
import { useRouter } from 'expo-router';
import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, EmptyState, IconButton, Screen, Section, Skeleton } from '@/core/ui';
import { formatRelativeDate } from '@/core/utils';
import {
  BODY_FIELD_META,
  deriveBmi,
  frequentlyLoggedFields,
  type BodyField,
  type WeightLogEntry,
} from '@/domain/body';

import { AddMeasurementsSheet } from '../components/AddMeasurementsSheet';
import { AddWeightSheet } from '../components/AddWeightSheet';
import { useBodyData } from '../hooks/useBodyData';
import { useBodyHeightCm } from '../hooks/useBodyHeightCm';

const COMPOSITION: readonly BodyField[] = ['weightKg', 'bodyFatPct', 'muscleMassKg', 'visceralFat'];
const CIRCUMFERENCES: readonly BodyField[] = [
  'neckCm',
  'chestCm',
  'waistCm',
  'hipsCm',
  'leftArmCm',
  'rightArmCm',
  'leftForearmCm',
  'rightForearmCm',
  'leftThighCm',
  'rightThighCm',
  'leftCalfCm',
  'rightCalfCm',
];

const fmt = (value: number): string => (value % 1 === 0 ? String(value) : value.toFixed(1));

/** Measurements home (UI_UX §4.4/§4.5) — current body state + the weight log. */
export function MeasurementsScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const data = useBodyData();
  const heightCm = useBodyHeightCm();
  const [weightOpen, setWeightOpen] = useState(false);
  const [measureOpen, setMeasureOpen] = useState(false);

  const lastWeight = data?.latest.weightKg?.value ?? null;
  const bmi = data?.latest.bmi?.value ?? deriveBmi(lastWeight, heightCm);
  const hasAny = data ? Object.values(data.latest).some((v) => v !== null) : false;

  const stat = (field: BodyField): ReactNode => {
    const latest = data?.latest[field] ?? null;
    if (latest === null) return null;
    const meta = BODY_FIELD_META[field];
    return (
      <View key={field} style={{ width: '50%', paddingVertical: theme.space.sm }}>
        <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>{meta.label}</Text>
        <Text
          style={{
            ...theme.type.heading,
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {fmt(latest.value)}
          <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>
            {meta.unit ? ` ${meta.unit}` : ''}
          </Text>
        </Text>
      </View>
    );
  };

  return (
    <Screen scroll>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          justifyContent: 'space-between',
          marginTop: theme.space.sm,
          marginBottom: theme.space.lg,
        }}
      >
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary }}>Measurements</Text>
        <IconButton
          icon={<Settings color={theme.color.textSecondary} size={24} strokeWidth={1.75} />}
          onPress={() => router.push('/settings')}
          accessibilityLabel="Settings"
        />
      </View>

      <View style={{ flexDirection: 'row', gap: theme.space.sm, marginBottom: theme.space.lg }}>
        <View style={{ flex: 1 }}>
          <Button label="Add Weight" onPress={() => setWeightOpen(true)} />
        </View>
        <View style={{ flex: 1 }}>
          <Button label="Measurements" variant="secondary" onPress={() => setMeasureOpen(true)} />
        </View>
      </View>

      {data === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={120} />
          <Skeleton height={120} />
        </View>
      ) : !hasAny ? (
        <EmptyState title="No measurements yet — tap Add Weight to start." />
      ) : (
        <View style={{ gap: theme.space.lg }}>
          <Section title="Current">
            <Card>
              <View style={{ flexDirection: 'row', flexWrap: 'wrap' }}>
                {COMPOSITION.map(stat)}
                {bmi !== null ? (
                  <View style={{ width: '50%', paddingVertical: theme.space.sm }}>
                    <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>
                      BMI{data.latest.bmi ? '' : ' (derived)'}
                    </Text>
                    <Text
                      style={{
                        ...theme.type.heading,
                        color: theme.color.textPrimary,
                        fontVariant: ['tabular-nums'],
                      }}
                    >
                      {bmi.toFixed(1)}
                    </Text>
                  </View>
                ) : null}
              </View>
            </Card>
          </Section>

          {CIRCUMFERENCES.some((f) => data.latest[f] !== null) ? (
            <Section title="Circumferences">
              <Card>
                <View style={{ flexDirection: 'row', flexWrap: 'wrap' }}>
                  {CIRCUMFERENCES.map(stat)}
                </View>
              </Card>
            </Section>
          ) : null}

          {data.weightLog.length > 0 ? (
            <Section title="Weight log">
              <Card>
                {data.weightLog.map((entry) => (
                  <WeightLogRow key={entry.date} entry={entry} />
                ))}
              </Card>
            </Section>
          ) : null}
        </View>
      )}

      <AddWeightSheet
        key={`w-${lastWeight ?? 'none'}`}
        visible={weightOpen}
        lastWeightKg={lastWeight}
        onClose={() => setWeightOpen(false)}
      />
      {data ? (
        <AddMeasurementsSheet
          visible={measureOpen}
          latest={data.latest}
          expanded={frequentlyLoggedFields(data.snapshots)}
          onClose={() => setMeasureOpen(false)}
        />
      ) : null}
    </Screen>
  );
}

function WeightLogRow(props: { readonly entry: WeightLogEntry }): ReactNode {
  const theme = useTheme();
  const { entry } = props;
  const down = entry.deltaKg !== null && entry.deltaKg < 0;
  return (
    <View
      style={{
        flexDirection: 'row',
        alignItems: 'center',
        justifyContent: 'space-between',
        paddingVertical: theme.space.sm,
      }}
    >
      <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
        {formatRelativeDate(entry.date)}
      </Text>
      <View style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.sm }}>
        {entry.deltaKg !== null && entry.deltaKg !== 0 ? (
          <View style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.xs }}>
            {down ? (
              <TrendingDown color={theme.color.textSecondary} size={16} />
            ) : (
              <TrendingUp color={theme.color.textSecondary} size={16} />
            )}
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              {Math.abs(entry.deltaKg).toFixed(1)} kg
            </Text>
          </View>
        ) : null}
        <Text
          style={{
            ...theme.type.bodyStrong,
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {entry.weightKg.toFixed(1)} kg
        </Text>
      </View>
    </View>
  );
}
