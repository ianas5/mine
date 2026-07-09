import { Trash2 } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { showToast } from '@/core/ui';
import { MEAL_SLOT_LABELS } from '@/domain/nutrition';
import type { MealEntry } from '@/domain/models';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';

const g = (value: number): string => (value % 1 === 0 ? String(value) : value.toFixed(1));

/**
 * One logged meal entry. Delete is immediate + a 5 s Undo toast (UI_UX §6 —
 * forgiving beats a confirm dialog for reversible actions). The macro figures are
 * the stored snapshot, not recomputed from the food.
 */
export function MealEntryRow(props: { readonly entry: MealEntry }): ReactNode {
  const theme = useTheme();
  const { entry } = props;

  const remove = (): void => {
    void nutritionRepository.deleteMealEntry(entry.id);
    showToast(`Removed ${entry.foodName}`, 'neutral', {
      label: 'Undo',
      onPress: () => void nutritionRepository.restoreMealEntry(entry),
    });
  };

  return (
    <View
      style={{
        flexDirection: 'row',
        alignItems: 'center',
        gap: theme.space.md,
        paddingVertical: theme.space.sm,
      }}
    >
      <View style={{ flex: 1 }}>
        <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>{entry.foodName}</Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
          {g(entry.loggedAmount)} {entry.loggedUnit}
          {entry.slot ? ` · ${MEAL_SLOT_LABELS[entry.slot]}` : ''}
        </Text>
      </View>
      <View style={{ alignItems: 'flex-end' }}>
        <Text
          style={{
            ...theme.type.bodyStrong,
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {entry.kcal} kcal
        </Text>
        <Text style={{ ...theme.type.micro, color: theme.color.textTertiary }}>
          {g(entry.proteinG)}P · {g(entry.carbG)}C · {g(entry.fatG)}F
        </Text>
      </View>
      <Pressable
        onPress={remove}
        accessibilityRole="button"
        accessibilityLabel={`Delete ${entry.foodName}`}
        hitSlop={theme.space.sm}
        style={({ pressed }) => ({
          width: 40,
          height: 40,
          alignItems: 'center',
          justifyContent: 'center',
          opacity: pressed ? 0.5 : 1,
        })}
      >
        <Trash2 color={theme.color.textTertiary} size={18} />
      </Pressable>
    </View>
  );
}
