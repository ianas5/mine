import { useRouter } from 'expo-router';
import { ArrowLeft, ChevronLeft, ChevronRight } from 'lucide-react-native';
import { useEffect, useState, type ReactNode } from 'react';
import { ScrollView, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { useTheme } from '@/core/theme';
import { Button, Card, IconButton, Input, showToast } from '@/core/ui';
import { addDaysIso, formatRelativeDate, todayIso } from '@/core/utils';
import type { NutritionTarget } from '@/domain/models';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';

import { useTargets } from '../hooks/useTargets';

const num = (s: string): number => {
  const n = Number.parseFloat(s);
  return Number.isFinite(n) ? n : 0;
};

/** Targets editor (UI_UX §4.7): "Set new targets from <date>", with past eras read-only. */
export function TargetsEditorScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const targets = useTargets();

  const [effectiveFrom, setEffectiveFrom] = useState(todayIso());
  const [kcal, setKcal] = useState('');
  const [protein, setProtein] = useState('');
  const [carb, setCarb] = useState('');
  const [fat, setFat] = useState('');
  const [water, setWater] = useState('');
  const [seeded, setSeeded] = useState(false);

  // Seed the form once from the target active today (convenience; editing stays a
  // *new era*, never a silent rewrite).
  useEffect(() => {
    if (seeded) return;
    let live = true;
    void nutritionRepository.resolveTargetForDate(todayIso()).then((t) => {
      if (!live) return;
      if (t) {
        setKcal(String(t.kcal));
        setProtein(String(t.proteinG));
        setCarb(String(t.carbG));
        setFat(String(t.fatG));
        setWater(t.waterMl !== null ? String(t.waterMl) : '');
      }
      setSeeded(true);
    });
    return () => {
      live = false;
    };
  }, [seeded]);

  const save = async (): Promise<void> => {
    await nutritionRepository.setTarget(effectiveFrom, {
      kcal: Math.round(num(kcal)),
      proteinG: num(protein),
      carbG: num(carb),
      fatG: num(fat),
      waterMl: water.trim() === '' ? null : Math.round(num(water)),
    });
    showToast('Targets saved', 'success');
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
          Targets
        </Text>
        <Button label="Save" size="md" onPress={() => void save()} />
      </View>

      <ScrollView
        contentContainerStyle={{
          paddingHorizontal: theme.space.lg,
          paddingBottom: insets.bottom + theme.space.xxxl,
          gap: theme.space.lg,
        }}
        keyboardShouldPersistTaps="handled"
      >
        <View style={{ gap: theme.space.sm }}>
          <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
            SET NEW TARGETS FROM
          </Text>
          <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
            <IconButton
              icon={<ChevronLeft color={theme.color.textPrimary} size={22} strokeWidth={1.75} />}
              onPress={() => setEffectiveFrom((d) => addDaysIso(d, -1))}
              accessibilityLabel="Earlier date"
            />
            <Text
              style={{
                ...theme.type.body,
                color: theme.color.textPrimary,
                flex: 1,
                textAlign: 'center',
              }}
            >
              {formatRelativeDate(effectiveFrom)} · {effectiveFrom}
            </Text>
            <IconButton
              icon={<ChevronRight color={theme.color.textPrimary} size={22} strokeWidth={1.75} />}
              onPress={() => setEffectiveFrom((d) => addDaysIso(d, 1))}
              accessibilityLabel="Later date"
            />
          </Card>
        </View>

        <NumField label="Calories (kcal)" value={kcal} onChange={setKcal} />
        <NumField label="Protein (g)" value={protein} onChange={setProtein} />
        <NumField label="Carbs (g)" value={carb} onChange={setCarb} />
        <NumField label="Fat (g)" value={fat} onChange={setFat} />
        <NumField label="Water (ml, optional)" value={water} onChange={setWater} />

        {targets && targets.length > 0 ? (
          <View style={{ gap: theme.space.sm, marginTop: theme.space.md }}>
            <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
              TARGET HISTORY
            </Text>
            {targets.map((t) => (
              <TargetEraRow key={t.id} target={t} />
            ))}
          </View>
        ) : null}
      </ScrollView>
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

function TargetEraRow(props: { readonly target: NutritionTarget }): ReactNode {
  const theme = useTheme();
  const { target } = props;
  return (
    <Card>
      <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
        From {formatRelativeDate(target.effectiveFrom)} · {target.effectiveFrom}
      </Text>
      <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
        {target.kcal.toLocaleString()} kcal · {target.proteinG}P {target.carbG}C {target.fatG}F
        {target.waterMl !== null ? ` · ${target.waterMl} ml water` : ''}
      </Text>
    </Card>
  );
}
