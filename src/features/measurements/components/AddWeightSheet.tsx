import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Sheet, Stepper, showToast } from '@/core/ui';
import { todayIso } from '@/core/utils';
import { bodyRepository } from '@/data/repositories/bodyRepository';

interface AddWeightSheetProps {
  readonly visible: boolean;
  readonly lastWeightKg: number | null;
  readonly onClose: () => void;
}

const DEFAULT_START_KG = 70;

/** Add Weight (UI_UX §4.4) — stepper pre-set to the last weight, 0.1 kg steps, Save. ≤ 3 taps. */
export function AddWeightSheet(props: AddWeightSheetProps): ReactNode {
  const theme = useTheme();
  // Seeded from the last weight; the call site keys this by baseline so a new
  // latest weigh-in re-seeds the stepper on the next open.
  const [weightKg, setWeightKg] = useState(props.lastWeightKg ?? DEFAULT_START_KG);

  const save = async (): Promise<void> => {
    await bodyRepository.saveSnapshot(todayIso(), { weightKg });
    showToast('Weight logged', 'success');
    props.onClose();
  };

  return (
    <Sheet visible={props.visible} onClose={props.onClose} title="Add Weight">
      <View style={{ gap: theme.space.xl, alignItems: 'center' }}>
        <View style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
          <Text
            style={{
              ...theme.type.display,
              color: theme.color.textPrimary,
              fontVariant: ['tabular-nums'],
            }}
          >
            {weightKg.toFixed(1)}
          </Text>
          <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>kg</Text>
        </View>
        <Stepper
          value={weightKg}
          onChange={setWeightKg}
          step={0.1}
          min={0}
          max={500}
          accessibilityLabel="Weight in kilograms"
        />
        <Button label="Save" onPress={() => void save()} />
      </View>
    </Sheet>
  );
}
