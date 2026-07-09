import { useLocalSearchParams, useRouter } from 'expo-router';
import { ChevronLeft, ChevronRight, SlidersHorizontal, UtensilsCrossed } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, EmptyState, IconButton, Screen, Skeleton } from '@/core/ui';
import { addDaysIso, formatRelativeDate, todayIso } from '@/core/utils';

import { LogMealSheet } from '../components/LogMealSheet';
import { MacroSummary } from '../components/MacroSummary';
import { MealEntryRow } from '../components/MealEntryRow';
import { WaterCard } from '../components/WaterCard';
import { useNutritionDay } from '../hooks/useNutritionDay';
import { useWaterCupMl } from '../hooks/useWaterCupMl';

/** Nutrition day view (UI_UX §4.3) — totals, the day's entries, and the Log Meal loop. */
export function NutritionScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const today = todayIso();
  const [date, setDate] = useState(today);
  const [logOpen, setLogOpen] = useState(false);
  const day = useNutritionDay(date);
  const cupMl = useWaterCupMl();

  // A dashboard quick action (?open=meal) drives the sheet open directly — no effect
  // (UI_UX §7.2, 1 tap to the sheet). Closing clears the intent so a repeat tap re-fires.
  const { open } = useLocalSearchParams<{ open?: string }>();
  const showLog = logOpen || open === 'meal';
  const closeLog = (): void => {
    setLogOpen(false);
    if (open === 'meal') router.setParams({ open: undefined });
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
        <IconButton
          icon={<ChevronLeft color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
          onPress={() => setDate((d) => addDaysIso(d, -1))}
          accessibilityLabel="Previous day"
        />
        <Pressable
          onPress={() => setDate(today)}
          accessibilityRole="button"
          accessibilityLabel="Jump to today"
        >
          <Text style={{ ...theme.type.title, color: theme.color.textPrimary }}>
            {formatRelativeDate(date)}
          </Text>
        </Pressable>
        <IconButton
          icon={<ChevronRight color={theme.color.textTertiary} size={24} strokeWidth={1.75} />}
          onPress={() => setDate((d) => (d < today ? addDaysIso(d, 1) : d))}
          accessibilityLabel="Next day"
        />
      </View>

      {day === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={120} />
          <Skeleton height={64} />
        </View>
      ) : (
        <View style={{ gap: theme.space.lg }}>
          <MacroSummary
            totals={day.totals}
            target={day.target}
            remaining={day.remaining}
            adherence={day.adherence}
          />

          <WaterCard
            date={date}
            waterMl={day.waterMl}
            targetMl={day.target?.waterMl ?? null}
            cupMl={cupMl}
          />

          <Button label="Log Meal" onPress={() => setLogOpen(true)} />

          {day.entries.length === 0 ? (
            <EmptyState title="Nothing logged yet — tap Log Meal to start." />
          ) : (
            <Card>
              {day.entries.map((entry) => (
                <MealEntryRow key={entry.id} entry={entry} />
              ))}
            </Card>
          )}

          <NavRow
            icon={<SlidersHorizontal color={theme.color.accent} size={22} strokeWidth={1.75} />}
            title="Targets"
            subtitle={
              day.target
                ? 'Update your daily goals from a date'
                : 'Set daily goals to see remaining'
            }
            onPress={() => router.push('/nutrition/targets')}
          />
          <NavRow
            icon={<UtensilsCrossed color={theme.color.accent} size={22} strokeWidth={1.75} />}
            title="Foods"
            subtitle="Create and edit reusable foods and quick meals"
            onPress={() => router.push('/nutrition/foods')}
          />
        </View>
      )}

      <LogMealSheet
        visible={showLog}
        date={date}
        nowHour={new Date().getHours()}
        onClose={closeLog}
        onCreateFood={() => {
          closeLog();
          router.push('/nutrition/foods/new');
        }}
      />
    </Screen>
  );
}

function NavRow(props: {
  readonly icon: ReactNode;
  readonly title: string;
  readonly subtitle: string;
  readonly onPress: () => void;
}): ReactNode {
  const theme = useTheme();
  return (
    <Pressable
      onPress={props.onPress}
      accessibilityRole="button"
      accessibilityLabel={props.title}
      style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
    >
      <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
        {props.icon}
        <View style={{ flex: 1 }}>
          <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>{props.title}</Text>
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            {props.subtitle}
          </Text>
        </View>
        <ChevronRight color={theme.color.textTertiary} size={20} />
      </Card>
    </Pressable>
  );
}
