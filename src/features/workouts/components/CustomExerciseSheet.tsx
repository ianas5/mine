import { zodResolver } from '@hookform/resolvers/zod';
import type { ReactNode } from 'react';
import { Controller, useForm } from 'react-hook-form';
import { ScrollView, Text, View } from 'react-native';

import { LOAD_TYPES, LOAD_TYPE_LABELS, MUSCLE_GROUPS, MUSCLE_GROUP_LABELS } from '@/domain/fitness';
import { useTheme } from '@/core/theme';
import { Button, Chip, Input, Sheet, showToast } from '@/core/ui';
import {
  DuplicateExerciseNameError,
  exerciseRepository,
} from '@/data/repositories/exerciseRepository';

import { customExerciseSchema, type CustomExerciseInput } from '../schemas/customExerciseSchema';

interface CustomExerciseSheetProps {
  readonly visible: boolean;
  readonly onClose: () => void;
}

const DEFAULTS: CustomExerciseInput = {
  name: '',
  primaryMuscleGroup: 'chest',
  loadType: 'external',
  defaultUnilateral: false,
};

/** Create-a-custom-exercise form — RHF + Zod (ARCHITECTURE §10, DESIGN_SYSTEM §6). */
export function CustomExerciseSheet(props: CustomExerciseSheetProps): ReactNode {
  const theme = useTheme();
  const { control, handleSubmit, reset, setError, formState } = useForm<CustomExerciseInput>({
    resolver: zodResolver(customExerciseSchema),
    defaultValues: DEFAULTS,
  });

  const close = (): void => {
    reset(DEFAULTS);
    props.onClose();
  };

  const onSubmit = handleSubmit(async (values) => {
    try {
      await exerciseRepository.createCustom(values);
      showToast(`Added ${values.name.trim()}`, 'success');
      close();
    } catch (cause) {
      if (cause instanceof DuplicateExerciseNameError) {
        setError('name', { message: cause.message });
        return;
      }
      showToast('Could not add exercise');
    }
  });

  const label = { ...theme.type.caption, color: theme.color.textSecondary } as const;

  return (
    <Sheet visible={props.visible} onClose={close} title="New Exercise" dirty={formState.isDirty}>
      <ScrollView keyboardShouldPersistTaps="handled" style={{ maxHeight: 480 }}>
        <View style={{ gap: theme.space.lg }}>
          <Controller
            control={control}
            name="name"
            render={({ field, fieldState }) => (
              <Input
                label="Name"
                value={field.value}
                onChangeText={field.onChange}
                onBlur={field.onBlur}
                placeholder="e.g. Cable Y-Raise"
                error={fieldState.error?.message}
                autoFocus
              />
            )}
          />

          <Controller
            control={control}
            name="primaryMuscleGroup"
            render={({ field }) => (
              <View style={{ gap: theme.space.sm }}>
                <Text style={label}>Primary muscle</Text>
                <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: theme.space.sm }}>
                  {MUSCLE_GROUPS.map((group) => (
                    <Chip
                      key={group}
                      label={MUSCLE_GROUP_LABELS[group]}
                      selected={field.value === group}
                      onPress={() => field.onChange(group)}
                    />
                  ))}
                </View>
              </View>
            )}
          />

          <Controller
            control={control}
            name="loadType"
            render={({ field }) => (
              <View style={{ gap: theme.space.sm }}>
                <Text style={label}>Load type</Text>
                <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: theme.space.sm }}>
                  {LOAD_TYPES.map((lt) => (
                    <Chip
                      key={lt}
                      label={LOAD_TYPE_LABELS[lt]}
                      selected={field.value === lt}
                      onPress={() => field.onChange(lt)}
                    />
                  ))}
                </View>
              </View>
            )}
          />

          <Controller
            control={control}
            name="defaultUnilateral"
            render={({ field }) => (
              <View style={{ gap: theme.space.sm }}>
                <Text style={label}>Unilateral (one side at a time)</Text>
                <View style={{ flexDirection: 'row', gap: theme.space.sm }}>
                  <Chip label="No" selected={!field.value} onPress={() => field.onChange(false)} />
                  <Chip label="Yes" selected={field.value} onPress={() => field.onChange(true)} />
                </View>
              </View>
            )}
          />

          <Button label="Add exercise" onPress={onSubmit} loading={formState.isSubmitting} />
        </View>
      </ScrollView>
    </Sheet>
  );
}
