import {
  ArrowLeftRight,
  ChevronRight,
  Images,
  Settings,
  TrendingDown,
  TrendingUp,
} from 'lucide-react-native';
import { useLocalSearchParams, useRouter } from 'expo-router';
import { useState, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

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

  // Dashboard quick actions (?open=weight | measure) drive the sheet open directly — no
  // effect (UI_UX §7.2, 1 tap). Closing clears the intent so a repeat tap re-fires.
  const { open } = useLocalSearchParams<{ open?: string }>();
  const showWeight = weightOpen || open === 'weight';
  const showMeasure = measureOpen || open === 'measure';
  const closeWeight = (): void => {
    setWeightOpen(false);
    if (open === 'weight') router.setParams({ open: undefined });
  };
  const closeMeasure = (): void => {
    setMeasureOpen(false);
    if (open === 'measure') router.setParams({ open: undefined });
  };

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

      <Pressable
        onPress={() => router.push('/measurements/photos')}
        accessibilityRole="button"
        accessibilityLabel="Progress photos"
        style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1, marginBottom: theme.space.lg })}
      >
        <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
          <Images color={theme.color.accent} size={22} strokeWidth={1.75} />
          <View style={{ flex: 1 }}>
            <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
              Progress photos
            </Text>
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              Capture and compare your visual progress
            </Text>
          </View>
          <ChevronRight color={theme.color.textTertiary} size={20} />
        </Card>
      </Pressable>

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

          {data.snapshots.length >= 2 ? (
            <Pressable
              onPress={() => router.push('/measurements/compare')}
              accessibilityRole="button"
              accessibilityLabel="Compare two dates"
              style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
            >
              <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
                <ArrowLeftRight color={theme.color.accent} size={22} strokeWidth={1.75} />
                <View style={{ flex: 1 }}>
                  <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
                    Compare
                  </Text>
                  <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
                    See the change between any two dates
                  </Text>
                </View>
                <ChevronRight color={theme.color.textTertiary} size={20} />
              </Card>
            </Pressable>
          ) : null}
        </View>
      )}

      <AddWeightSheet
        key={`w-${lastWeight ?? 'none'}`}
        visible={showWeight}
        lastWeightKg={lastWeight}
        onClose={closeWeight}
      />
      {data ? (
        <AddMeasurementsSheet
          visible={showMeasure}
          latest={data.latest}
          expanded={frequentlyLoggedFields(data.snapshots)}
          onClose={closeMeasure}
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
