import { Plus } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, ScrollView, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Chip, Input, Sheet, Stepper, showToast } from '@/core/ui';
import {
  MEAL_SLOTS,
  MEAL_SLOT_LABELS,
  defaultSlotForHour,
  scalePortion,
  type MealSlot,
} from '@/domain/nutrition';
import { nutritionRepository, type FoodPick } from '@/data/repositories/nutritionRepository';

import { useFoodPicks } from '../hooks/useFoodPicks';

interface LogMealSheetProps {
  readonly visible: boolean;
  readonly date: string;
  readonly nowHour: number;
  readonly onClose: () => void;
  readonly onCreateFood: () => void;
}

/**
 * The Log Meal sheet (UI_UX §4.3) — Recent & Frequent first, tap a food to a
 * portion pre-filled with its last-used amount and most-frequent slot, then Save.
 * Repeated meals are ≤ 3 taps (open · food · Save).
 */
export function LogMealSheet(props: LogMealSheetProps): ReactNode {
  const theme = useTheme();
  const picks = useFoodPicks();
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState<FoodPick | null>(null);
  const [amount, setAmount] = useState(0);
  const [slot, setSlot] = useState<MealSlot>('lunch');

  const reset = (): void => {
    setSelected(null);
    setQuery('');
  };

  const close = (): void => {
    reset();
    props.onClose();
  };

  const choose = (pick: FoodPick): void => {
    setSelected(pick);
    setAmount(pick.lastAmount ?? pick.food.servingAmount);
    setSlot(pick.slot ?? defaultSlotForHour(props.nowHour));
  };

  const save = async (): Promise<void> => {
    if (!selected) return;
    const macros = scalePortion(selected.food, amount);
    await nutritionRepository.addMealEntry({
      date: props.date,
      slot,
      foodId: selected.food.id,
      foodName: selected.food.name,
      loggedAmount: amount,
      loggedUnit: selected.food.servingUnit,
      ...macros,
    });
    showToast(`Logged ${selected.food.name}`, 'success');
    close();
  };

  const filtered = (picks ?? []).filter((p) =>
    p.food.name.toLowerCase().includes(query.trim().toLowerCase()),
  );

  const unitStep = (unit: string): number => (unit === 'g' || unit === 'ml' ? 5 : 0.5);

  return (
    <Sheet visible={props.visible} onClose={close} title="Log Meal">
      {selected ? (
        <View style={{ gap: theme.space.lg }}>
          <View>
            <Text style={{ ...theme.type.heading, color: theme.color.textPrimary }}>
              {selected.food.name}
            </Text>
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              per {selected.food.servingAmount} {selected.food.servingUnit} · {selected.food.kcal}{' '}
              kcal
            </Text>
          </View>

          <View
            style={{ flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between' }}
          >
            <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
              Amount ({selected.food.servingUnit})
            </Text>
            <Stepper
              value={amount}
              onChange={setAmount}
              step={unitStep(selected.food.servingUnit)}
              min={0}
              max={5000}
              accessibilityLabel="Portion amount"
            />
          </View>

          <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: theme.space.xs }}>
            {MEAL_SLOTS.map((s) => (
              <Chip
                key={s}
                label={MEAL_SLOT_LABELS[s]}
                selected={slot === s}
                onPress={() => setSlot(s)}
              />
            ))}
          </View>

          <Text
            style={{
              ...theme.type.body,
              color: theme.color.accent,
              fontVariant: ['tabular-nums'],
              textAlign: 'center',
            }}
          >
            {(() => {
              const m = scalePortion(selected.food, amount);
              return `${m.kcal} kcal · ${m.proteinG}P · ${m.carbG}C · ${m.fatG}F`;
            })()}
          </Text>

          <Button label="Save" onPress={() => void save()} />
          <Button label="Back" variant="ghost" size="md" onPress={() => setSelected(null)} />
        </View>
      ) : (
        <View style={{ gap: theme.space.md }}>
          <Input
            value={query}
            onChangeText={setQuery}
            placeholder="Search foods…"
            accessibilityLabel="Search foods"
          />
          <Pressable
            onPress={props.onCreateFood}
            accessibilityRole="button"
            accessibilityLabel="New food"
            style={({ pressed }) => ({
              flexDirection: 'row',
              alignItems: 'center',
              gap: theme.space.sm,
              paddingVertical: theme.space.sm,
              opacity: pressed ? 0.6 : 1,
            })}
          >
            <Plus color={theme.color.accent} size={18} />
            <Text style={{ ...theme.type.bodyStrong, color: theme.color.accent }}>New food</Text>
          </Pressable>

          <ScrollView style={{ maxHeight: 360 }} keyboardShouldPersistTaps="handled">
            {filtered.map((pick) => (
              <Pressable
                key={pick.food.id}
                onPress={() => choose(pick)}
                accessibilityRole="button"
                accessibilityLabel={pick.food.name}
                style={({ pressed }) => ({
                  flexDirection: 'row',
                  alignItems: 'center',
                  gap: theme.space.md,
                  paddingVertical: theme.space.md,
                  opacity: pressed ? 0.6 : 1,
                })}
              >
                <View style={{ flex: 1 }}>
                  <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
                    {pick.food.name}
                    {pick.food.isQuickMeal ? '  ·  quick meal' : ''}
                  </Text>
                  <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
                    {pick.food.kcal} kcal · {pick.food.proteinG}P per {pick.food.servingAmount}{' '}
                    {pick.food.servingUnit}
                  </Text>
                </View>
              </Pressable>
            ))}
          </ScrollView>
        </View>
      )}
    </Sheet>
  );
}
