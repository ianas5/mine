import { useLocalSearchParams, useRouter } from 'expo-router';
import { ArrowLeft } from 'lucide-react-native';
import { useEffect, useState, type ReactNode } from 'react';
import { ScrollView, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { useTheme } from '@/core/theme';
import { Button, Chip, Dialog, IconButton, Input, Skeleton, showToast } from '@/core/ui';
import {
  SERVING_UNITS,
  atwaterKcal,
  isKcalImplausible,
  type ServingUnit,
} from '@/domain/nutrition';
import { nutritionRepository, type FoodInput } from '@/data/repositories/nutritionRepository';

const num = (s: string): number => {
  const n = Number.parseFloat(s);
  return Number.isFinite(n) ? n : 0;
};

/** Create or edit a reusable food / quick meal (Phase 9). `id = 'new'` is create mode. */
export function FoodEditorScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const params = useLocalSearchParams<{ id: string }>();
  const id = params.id ?? 'new';
  const isNew = id === 'new';

  const [loaded, setLoaded] = useState(isNew);
  const [name, setName] = useState('');
  const [servingAmount, setServingAmount] = useState('100');
  const [unit, setUnit] = useState<ServingUnit>('g');
  const [kcal, setKcal] = useState('');
  const [protein, setProtein] = useState('');
  const [carb, setCarb] = useState('');
  const [fat, setFat] = useState('');
  const [quick, setQuick] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);

  useEffect(() => {
    if (isNew) return;
    let live = true;
    void nutritionRepository.getFood(id).then((food) => {
      if (!live || !food) return;
      setName(food.name);
      setServingAmount(String(food.servingAmount));
      setUnit(food.servingUnit);
      setKcal(String(food.kcal));
      setProtein(String(food.proteinG));
      setCarb(String(food.carbG));
      setFat(String(food.fatG));
      setQuick(food.isQuickMeal);
      setLoaded(true);
    });
    return () => {
      live = false;
    };
  }, [id, isNew]);

  const macros = { kcal: num(kcal), proteinG: num(protein), carbG: num(carb), fatG: num(fat) };
  const implausible = isKcalImplausible(macros);

  const save = async (): Promise<void> => {
    const input: FoodInput = {
      name,
      servingAmount: num(servingAmount) || 1,
      servingUnit: unit,
      kcal: Math.round(macros.kcal),
      proteinG: macros.proteinG,
      carbG: macros.carbG,
      fatG: macros.fatG,
      isQuickMeal: quick,
    };
    if (isNew) await nutritionRepository.createFood(input);
    else await nutritionRepository.updateFood(id, input);
    showToast('Food saved', 'success');
    router.back();
  };

  const remove = async (): Promise<void> => {
    setConfirmDelete(false);
    await nutritionRepository.deleteFood(id);
    showToast('Food deleted');
    router.back();
  };

  return (
    <View style={{ flex: 1, backgroundColor: theme.color.bg, paddingTop: insets.top }}>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.sm,
          paddingHorizontal: theme.space.lg,
          paddingVertical: theme.space.sm,
        }}
      >
        <IconButton
          icon={<ArrowLeft color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
          onPress={() => router.back()}
          accessibilityLabel="Back"
        />
        <Text style={{ ...theme.type.heading, color: theme.color.textPrimary, flex: 1 }}>
          {isNew ? 'New food' : 'Edit food'}
        </Text>
        <Button label="Save" size="md" onPress={() => void save()} />
      </View>

      {!loaded ? (
        <View style={{ padding: theme.space.lg, gap: theme.space.sm }}>
          <Skeleton height={48} />
          <Skeleton height={120} />
        </View>
      ) : (
        <ScrollView
          contentContainerStyle={{
            paddingHorizontal: theme.space.lg,
            paddingBottom: insets.bottom + theme.space.xxxl,
            gap: theme.space.lg,
          }}
          keyboardShouldPersistTaps="handled"
        >
          <Input
            value={name}
            onChangeText={setName}
            placeholder="Food name"
            accessibilityLabel="Food name"
          />

          <View style={{ flexDirection: 'row', gap: theme.space.md }}>
            <View style={{ flex: 1 }}>
              <Input
                value={servingAmount}
                onChangeText={setServingAmount}
                keyboardType="decimal-pad"
                label="Serving"
                accessibilityLabel="Serving amount"
              />
            </View>
          </View>
          <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: theme.space.xs }}>
            {SERVING_UNITS.map((u) => (
              <Chip key={u} label={u} selected={unit === u} onPress={() => setUnit(u)} />
            ))}
          </View>

          <NumField label="Calories (kcal)" value={kcal} onChange={setKcal} />
          <NumField label="Protein (g)" value={protein} onChange={setProtein} />
          <NumField label="Carbs (g)" value={carb} onChange={setCarb} />
          <NumField label="Fat (g)" value={fat} onChange={setFat} />

          {implausible ? (
            <Text style={{ ...theme.type.caption, color: theme.color.attention }}>
              Calories look off versus the macros (~{Math.round(atwaterKcal(macros))} kcal from
              protein/carbs/fat). Saved as entered — double-check if this wasn’t intentional.
            </Text>
          ) : null}

          <Chip
            label={quick ? 'Quick meal ✓' : 'Mark as quick meal'}
            selected={quick}
            onPress={() => setQuick((q) => !q)}
          />

          {!isNew ? (
            <Button
              label="Delete food"
              variant="destructive"
              size="md"
              onPress={() => setConfirmDelete(true)}
            />
          ) : null}
        </ScrollView>
      )}

      <Dialog
        visible={confirmDelete}
        title="Delete food?"
        message="Past meal entries keep their snapshot but lose the link to this food."
        confirmLabel="Delete"
        onConfirm={() => void remove()}
        onCancel={() => setConfirmDelete(false)}
      />
    </View>
  );
}

function NumField(props: {
  readonly label: string;
  readonly value: string;
  readonly onChange: (next: string) => void;
}): ReactNode {
  return (
    <Input
      value={props.value}
      onChangeText={props.onChange}
      keyboardType="decimal-pad"
      label={props.label}
      placeholder="0"
      accessibilityLabel={props.label}
    />
  );
}
