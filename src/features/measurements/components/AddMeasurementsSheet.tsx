import { useState, type ReactNode } from 'react';
import { Pressable, ScrollView, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Input, Sheet, showToast } from '@/core/ui';
import { todayIso } from '@/core/utils';
import { BODY_FIELDS, BODY_FIELD_META, type BodyField, type FieldLatest } from '@/domain/body';
import { bodyRepository, type MeasurementPatch } from '@/data/repositories/bodyRepository';

interface AddMeasurementsSheetProps {
  readonly visible: boolean;
  readonly latest: Record<BodyField, FieldLatest | null>;
  readonly expanded: ReadonlySet<BodyField>;
  readonly onClose: () => void;
}

const num = (s: string | undefined): number | undefined => {
  if (s === undefined || s.trim() === '') return undefined;
  const n = Number.parseFloat(s);
  return Number.isFinite(n) ? n : undefined;
};

/**
 * Add Measurements (UI_UX §4.5) — every site, last value as placeholder, fill any
 * subset. Frequently co-logged fields start expanded; the rest sit behind "More
 * sites". Saving merge-upserts only the filled fields — omissions never clear.
 */
export function AddMeasurementsSheet(props: AddMeasurementsSheetProps): ReactNode {
  const theme = useTheme();
  const [values, setValues] = useState<Partial<Record<BodyField, string>>>({});
  const [showMore, setShowMore] = useState(false);

  const set = (field: BodyField, text: string): void => setValues((v) => ({ ...v, [field]: text }));

  const close = (): void => {
    setValues({});
    setShowMore(false);
    props.onClose();
  };

  const save = async (): Promise<void> => {
    const patch: MeasurementPatch = {};
    for (const field of BODY_FIELDS) {
      const value = num(values[field]);
      if (value !== undefined) patch[field] = value;
    }
    if (Object.keys(patch).length === 0) {
      close();
      return;
    }
    await bodyRepository.saveSnapshot(todayIso(), patch);
    showToast('Measurements saved', 'success');
    close();
  };

  const hiddenCount = BODY_FIELDS.filter((f) => !props.expanded.has(f)).length;

  const fieldRow = (field: BodyField): ReactNode => {
    const meta = BODY_FIELD_META[field];
    const last = props.latest[field];
    return (
      <Input
        key={field}
        value={values[field] ?? ''}
        onChangeText={(t) => set(field, t)}
        keyboardType="decimal-pad"
        returnKeyType="next"
        label={`${meta.label}${meta.unit ? ` (${meta.unit})` : ''}`}
        placeholder={last ? `Last: ${last.value}` : '—'}
        accessibilityLabel={meta.label}
      />
    );
  };

  return (
    <Sheet visible={props.visible} onClose={close} title="Add Measurements">
      <View style={{ gap: theme.space.md }}>
        <ScrollView style={{ maxHeight: 420 }} keyboardShouldPersistTaps="handled">
          <View style={{ gap: theme.space.md }}>
            {BODY_FIELDS.filter((f) => props.expanded.has(f)).map(fieldRow)}

            {showMore ? (
              <>{BODY_FIELDS.filter((f) => !props.expanded.has(f)).map(fieldRow)}</>
            ) : hiddenCount > 0 ? (
              <Pressable
                onPress={() => setShowMore(true)}
                accessibilityRole="button"
                accessibilityLabel="More sites"
                style={({ pressed }) => ({
                  paddingVertical: theme.space.sm,
                  opacity: pressed ? 0.6 : 1,
                })}
              >
                <Text style={{ ...theme.type.bodyStrong, color: theme.color.accent }}>
                  More sites ({hiddenCount})
                </Text>
              </Pressable>
            ) : null}
          </View>
        </ScrollView>

        <Button label="Save" onPress={() => void save()} />
      </View>
    </Sheet>
  );
}
