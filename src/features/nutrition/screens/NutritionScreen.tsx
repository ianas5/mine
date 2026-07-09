import { useRouter } from 'expo-router';
import { ChevronLeft, ChevronRight, UtensilsCrossed } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, EmptyState, IconButton, Screen, Skeleton } from '@/core/ui';
import { addDaysIso, formatRelativeDate, todayIso } from '@/core/utils';

import { LogMealSheet } from '../components/LogMealSheet';
import { MacroSummary } from '../components/MacroSummary';
import { MealEntryRow } from '../components/MealEntryRow';
import { useNutritionDay } from '../hooks/useNutritionDay';

/** Nutrition day view (UI_UX §4.3) — totals, the day's entries, and the Log Meal loop. */
export function NutritionScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const today = todayIso();
  const [date, setDate] = useState(today);
  const [logOpen, setLogOpen] = useState(false);
  const day = useNutritionDay(date);

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
          <MacroSummary totals={day.totals} />

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

          <Pressable
            onPress={() => router.push('/nutrition/foods')}
            accessibilityRole="button"
            accessibilityLabel="Manage foods"
            style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
          >
            <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
              <UtensilsCrossed color={theme.color.accent} size={22} strokeWidth={1.75} />
              <View style={{ flex: 1 }}>
                <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>Foods</Text>
                <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
                  Create and edit reusable foods and quick meals
                </Text>
              </View>
              <ChevronRight color={theme.color.textTertiary} size={20} />
            </Card>
          </Pressable>
        </View>
      )}

      <LogMealSheet
        visible={logOpen}
        date={date}
        nowHour={new Date().getHours()}
        onClose={() => setLogOpen(false)}
        onCreateFood={() => {
          setLogOpen(false);
          router.push('/nutrition/foods/new');
        }}
      />
    </Screen>
  );
}
