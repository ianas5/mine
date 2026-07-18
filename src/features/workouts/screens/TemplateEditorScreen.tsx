import { useLocalSearchParams, useRouter } from 'expo-router';
import { ArrowLeft, Trash2 } from 'lucide-react-native';
import { useEffect, useState, type ReactNode } from 'react';
import { ScrollView, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { useTheme } from '@/core/theme';
import { Button, Card, Chip, IconButton, Input, Skeleton, Stepper, showToast } from '@/core/ui';
import { weekdayLabel } from '@/core/utils';
import type { Exercise } from '@/domain/models';
import {
  programRepository,
  type TemplateExerciseInput,
  type TemplateInput,
} from '@/data/repositories/programRepository';

import { ExercisePickerSheet } from '../components/ExercisePickerSheet';

/** A locally-edited exercise line; 0 means "no target" and is saved as null. */
interface EditingExercise {
  readonly exerciseId: string;
  readonly name: string;
  sets: number;
  repMin: number;
  repMax: number;
  rpe: number;
  restSeconds: number;
}

const nullIfZero = (value: number): number | null => (value === 0 ? null : value);

function toInput(name: string, weekdays: number[], list: EditingExercise[]): TemplateInput {
  return {
    name,
    weekdays,
    notes: null,
    exercises: list.map((e): TemplateExerciseInput => ({
      exerciseId: e.exerciseId,
      targetSets: nullIfZero(e.sets),
      targetRepMin: nullIfZero(e.repMin),
      targetRepMax: nullIfZero(e.repMax),
      targetRpe: nullIfZero(e.rpe),
      restSeconds: nullIfZero(e.restSeconds),
      notes: null,
    })),
  };
}

/** Edit one session template — its name, weekday, and exercises with targets (Phase 8). */
export function TemplateEditorScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const insets = useSafeAreaInsets();
  const params = useLocalSearchParams<{ id: string }>();
  const templateId = params.id ?? '';

  const [loaded, setLoaded] = useState(false);
  const [name, setName] = useState('');
  const [weekdays, setWeekdays] = useState<number[]>([]);
  const [exercises, setExercises] = useState<EditingExercise[]>([]);
  const [pickerOpen, setPickerOpen] = useState(false);

  useEffect(() => {
    let live = true;
    void programRepository.getTemplate(templateId).then((template) => {
      if (!live || !template) return;
      setName(template.name);
      setWeekdays([...template.weekdays]);
      setExercises(
        template.exercises.map((te) => ({
          exerciseId: te.exerciseId,
          name: te.name,
          sets: te.target.sets ?? 0,
          repMin: te.target.repMin ?? 0,
          repMax: te.target.repMax ?? 0,
          rpe: te.target.rpe ?? 0,
          restSeconds: te.target.restSeconds ?? 0,
        })),
      );
      setLoaded(true);
    });
    return () => {
      live = false;
    };
  }, [templateId]);

  const patch = (index: number, next: Partial<EditingExercise>): void =>
    setExercises((list) => list.map((e, i) => (i === index ? { ...e, ...next } : e)));

  const addExercise = (exercise: Exercise): void =>
    setExercises((list) => [
      ...list,
      {
        exerciseId: exercise.id,
        name: exercise.name,
        sets: 3,
        repMin: 8,
        repMax: 12,
        rpe: 0,
        restSeconds: 0,
      },
    ]);

  const toggleWeekday = (day: number): void =>
    setWeekdays((list) =>
      list.includes(day) ? list.filter((d) => d !== day) : [...list, day].sort((a, b) => a - b),
    );

  const save = async (): Promise<void> => {
    await programRepository.updateTemplate(templateId, toInput(name, weekdays, exercises));
    showToast('Template saved', 'success');
    router.back();
  };

  const remove = async (): Promise<void> => {
    await programRepository.deleteTemplate(templateId);
    showToast('Template deleted');
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
          Edit session
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
            placeholder="Session name (e.g. Push A)"
            accessibilityLabel="Session name"
          />

          <View style={{ gap: theme.space.sm }}>
            <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
              WEEKDAYS · REPEATS ON EACH SELECTED DAY
            </Text>
            <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: theme.space.xs }}>
              {[0, 1, 2, 3, 4, 5, 6].map((d) => (
                <Chip
                  key={d}
                  label={weekdayLabel(d).slice(0, 3)}
                  selected={weekdays.includes(d)}
                  onPress={() => toggleWeekday(d)}
                />
              ))}
            </View>
          </View>

          {exercises.map((ex, index) => (
            <Card key={`${ex.exerciseId}-${index}`} style={{ gap: theme.space.md }}>
              <View
                style={{
                  flexDirection: 'row',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                }}
              >
                <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary, flex: 1 }}>
                  {ex.name}
                </Text>
                <IconButton
                  icon={<Trash2 color={theme.color.textTertiary} size={20} />}
                  onPress={() => setExercises((list) => list.filter((_, i) => i !== index))}
                  accessibilityLabel={`Remove ${ex.name}`}
                />
              </View>

              <TargetRow
                label="Sets"
                value={ex.sets}
                onChange={(v) => patch(index, { sets: v })}
                max={12}
              />
              <TargetRow
                label="Rep min"
                value={ex.repMin}
                onChange={(v) => patch(index, { repMin: v })}
                max={50}
              />
              <TargetRow
                label="Rep max"
                value={ex.repMax}
                onChange={(v) => patch(index, { repMax: v })}
                max={50}
              />
              <TargetRow
                label="RPE"
                value={ex.rpe}
                onChange={(v) => patch(index, { rpe: v })}
                step={0.5}
                max={10}
              />
              <TargetRow
                label="Rest (s)"
                value={ex.restSeconds}
                onChange={(v) => patch(index, { restSeconds: v })}
                step={15}
                max={600}
              />
            </Card>
          ))}

          <Button label="Add exercise" variant="secondary" onPress={() => setPickerOpen(true)} />

          <Button
            label="Delete session"
            variant="destructive"
            size="md"
            onPress={() => void remove()}
          />
        </ScrollView>
      )}

      <ExercisePickerSheet
        visible={pickerOpen}
        onClose={() => setPickerOpen(false)}
        onPick={(exercise) => addExercise(exercise)}
      />
    </View>
  );
}

function TargetRow(props: {
  readonly label: string;
  readonly value: number;
  readonly onChange: (next: number) => void;
  readonly step?: number;
  readonly max: number;
}): ReactNode {
  const theme = useTheme();
  return (
    <View style={{ flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between' }}>
      <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>{props.label}</Text>
      <Stepper
        value={props.value}
        onChange={props.onChange}
        step={props.step ?? 1}
        min={0}
        max={props.max}
        accessibilityLabel={props.label}
      />
    </View>
  );
}
