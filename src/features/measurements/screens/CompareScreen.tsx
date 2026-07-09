import { useRouter } from 'expo-router';
import { ArrowLeft, ArrowRight, Minus, TrendingDown, TrendingUp } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme, type Theme } from '@/core/theme';
import { Card, EmptyState, IconButton, Screen, Sheet, Skeleton } from '@/core/ui';
import { formatRelativeDate } from '@/core/utils';
import {
  BODY_FIELD_META,
  compareSnapshots,
  type ChangeDirection,
  type FieldComparison,
} from '@/domain/body';

import { useBodyData } from '../hooks/useBodyData';
import { useBodyHeightCm } from '../hooks/useBodyHeightCm';

const fmt = (value: number | null): string =>
  value === null ? '—' : value % 1 === 0 ? String(value) : value.toFixed(1);

function directionColor(theme: Theme, direction: ChangeDirection): string {
  switch (direction) {
    case 'improving':
      return theme.color.positive;
    case 'declining':
      return theme.color.danger;
    case 'incomparable':
      return theme.color.textTertiary;
    default:
      return theme.color.textSecondary;
  }
}

/** Compare any two body snapshots (FITNESS_DOMAIN §5.4). A field on only one date shows "—". */
export function CompareScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const data = useBodyData();
  const heightCm = useBodyHeightCm();
  const [pickA, setPickA] = useState<string | null>(null);
  const [pickB, setPickB] = useState<string | null>(null);
  const [picking, setPicking] = useState<null | 'a' | 'b'>(null);

  const dates = data?.snapshots.map((s) => s.date) ?? [];
  const dateB = pickB ?? dates[0] ?? null;
  const dateA = pickA ?? dates[1] ?? dates[0] ?? null;

  const snapshotA = data?.snapshots.find((s) => s.date === dateA) ?? null;
  const snapshotB = data?.snapshots.find((s) => s.date === dateB) ?? null;

  const rows =
    snapshotA && snapshotB
      ? compareSnapshots(snapshotA, snapshotB, heightCm).filter((c) => c.a !== null || c.b !== null)
      : [];

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
          Compare
        </Text>
      </View>

      {data === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={72} />
          <Skeleton height={160} />
        </View>
      ) : dates.length < 2 ? (
        <EmptyState title="Log at least two measurement days to compare them." />
      ) : (
        <View style={{ gap: theme.space.lg }}>
          <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.sm }}>
            <DateButton label={dateA} onPress={() => setPicking('a')} />
            <ArrowRight color={theme.color.textTertiary} size={20} />
            <DateButton label={dateB} onPress={() => setPicking('b')} />
          </Card>

          <Card>
            <View
              style={{
                flexDirection: 'row',
                paddingBottom: theme.space.sm,
                borderBottomWidth: 1,
                borderBottomColor: theme.color.border,
              }}
            >
              <Text style={{ ...theme.type.micro, color: theme.color.textTertiary, flex: 1.4 }}>
                METRIC
              </Text>
              <Text
                style={{
                  ...theme.type.micro,
                  color: theme.color.textTertiary,
                  flex: 1,
                  textAlign: 'right',
                }}
              >
                A
              </Text>
              <Text
                style={{
                  ...theme.type.micro,
                  color: theme.color.textTertiary,
                  flex: 1,
                  textAlign: 'right',
                }}
              >
                B
              </Text>
              <Text
                style={{
                  ...theme.type.micro,
                  color: theme.color.textTertiary,
                  flex: 1.3,
                  textAlign: 'right',
                }}
              >
                CHANGE
              </Text>
            </View>
            {rows.map((row) => (
              <CompareRow key={row.field} row={row} />
            ))}
          </Card>
        </View>
      )}

      <Sheet
        visible={picking !== null}
        onClose={() => setPicking(null)}
        title={picking === 'a' ? 'Date A' : 'Date B'}
      >
        <View>
          {dates.map((date) => (
            <Pressable
              key={date}
              onPress={() => {
                if (picking === 'a') setPickA(date);
                else setPickB(date);
                setPicking(null);
              }}
              accessibilityRole="button"
              accessibilityLabel={date}
              style={({ pressed }) => ({
                paddingVertical: theme.space.md,
                opacity: pressed ? 0.6 : 1,
              })}
            >
              <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
                {formatRelativeDate(date)} · {date}
              </Text>
            </Pressable>
          ))}
        </View>
      </Sheet>
    </Screen>
  );
}

function DateButton(props: {
  readonly label: string | null;
  readonly onPress: () => void;
}): ReactNode {
  const theme = useTheme();
  return (
    <Pressable
      onPress={props.onPress}
      accessibilityRole="button"
      accessibilityLabel={`Choose date (${props.label ?? 'none'})`}
      style={({ pressed }) => ({
        flex: 1,
        paddingVertical: theme.space.sm,
        paddingHorizontal: theme.space.md,
        borderRadius: theme.radius.md,
        borderWidth: 1,
        borderColor: theme.color.border,
        opacity: pressed ? 0.6 : 1,
      })}
    >
      <Text style={{ ...theme.type.caption, color: theme.color.textPrimary, textAlign: 'center' }}>
        {props.label ? formatRelativeDate(props.label) : '—'}
      </Text>
    </Pressable>
  );
}

function CompareRow(props: { readonly row: FieldComparison }): ReactNode {
  const theme = useTheme();
  const { row } = props;
  const meta = BODY_FIELD_META[row.field];
  const color = directionColor(theme, row.direction);
  const showChange = row.deltaAbs !== null && row.direction !== 'incomparable';

  return (
    <View style={{ flexDirection: 'row', alignItems: 'center', paddingVertical: theme.space.sm }}>
      <Text style={{ ...theme.type.caption, color: theme.color.textSecondary, flex: 1.4 }}>
        {meta.label}
      </Text>
      <Text
        style={{
          ...theme.type.caption,
          color: theme.color.textPrimary,
          flex: 1,
          textAlign: 'right',
          fontVariant: ['tabular-nums'],
        }}
      >
        {fmt(row.a)}
      </Text>
      <Text
        style={{
          ...theme.type.caption,
          color: theme.color.textPrimary,
          flex: 1,
          textAlign: 'right',
          fontVariant: ['tabular-nums'],
        }}
      >
        {fmt(row.b)}
      </Text>
      <View
        style={{
          flex: 1.3,
          flexDirection: 'row',
          alignItems: 'center',
          justifyContent: 'flex-end',
          gap: theme.space.xs,
        }}
      >
        {showChange ? (
          <>
            {row.direction === 'stable' || row.deltaAbs === 0 ? (
              <Minus color={color} size={14} />
            ) : (row.deltaAbs ?? 0) < 0 ? (
              <TrendingDown color={color} size={14} />
            ) : (
              <TrendingUp color={color} size={14} />
            )}
            <Text style={{ ...theme.type.caption, color, fontVariant: ['tabular-nums'] }}>
              {(row.deltaAbs ?? 0) > 0 ? '+' : ''}
              {fmt(row.deltaAbs)}
              {row.deltaPct !== null ? ` (${row.deltaPct > 0 ? '+' : ''}${row.deltaPct}%)` : ''}
            </Text>
          </>
        ) : (
          <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>—</Text>
        )}
      </View>
    </View>
  );
}
