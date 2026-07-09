import { useRouter } from 'expo-router';
import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { LOAD_TYPE_LABELS, MUSCLE_GROUP_LABELS } from '@/domain/fitness';
import { useTheme } from '@/core/theme';
import { Button, Dialog, Sheet, showToast } from '@/core/ui';
import { exerciseRepository } from '@/data/repositories/exerciseRepository';
import type { Exercise } from '@/domain/models';

interface ExerciseActionsSheetProps {
  readonly exercise: Exercise | null;
  readonly onClose: () => void;
}

/** Per-exercise archive / unarchive / delete actions (Phase 3 archive flow). */
export function ExerciseActionsSheet(props: ExerciseActionsSheetProps): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const [confirmingDelete, setConfirmingDelete] = useState(false);
  const exercise = props.exercise;

  const openReport = (): void => {
    if (!exercise) return;
    const id = exercise.id;
    props.onClose();
    router.push(`/workouts/exercise/${id}`);
  };

  const runAndClose = async (action: Promise<void>, message: string): Promise<void> => {
    await action;
    showToast(message);
    props.onClose();
  };

  return (
    <>
      <Sheet visible={exercise !== null} onClose={props.onClose} title={exercise?.name}>
        {exercise ? (
          <View style={{ gap: theme.space.lg }}>
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              {MUSCLE_GROUP_LABELS[exercise.primaryMuscleGroup]} ·{' '}
              {LOAD_TYPE_LABELS[exercise.loadType]}
              {exercise.defaultUnilateral ? ' · unilateral' : ''}
              {exercise.isCustom ? ' · custom' : ''}
            </Text>

            <Button label="View report" onPress={openReport} />

            {exercise.isArchived ? (
              <Button
                label="Unarchive"
                variant="secondary"
                onPress={() =>
                  void runAndClose(exerciseRepository.unarchive(exercise.id), 'Unarchived')
                }
              />
            ) : (
              <Button
                label="Archive"
                variant="secondary"
                onPress={() =>
                  void runAndClose(exerciseRepository.archive(exercise.id), 'Archived')
                }
              />
            )}

            {exercise.isCustom ? (
              <Button
                label="Delete permanently"
                variant="destructive"
                onPress={() => setConfirmingDelete(true)}
              />
            ) : null}
          </View>
        ) : null}
      </Sheet>

      <Dialog
        visible={confirmingDelete}
        title="Delete exercise?"
        message="This custom exercise will be permanently removed."
        confirmLabel="Delete"
        onConfirm={() => {
          setConfirmingDelete(false);
          if (exercise) {
            void runAndClose(exerciseRepository.remove(exercise.id), 'Deleted');
          }
        }}
        onCancel={() => setConfirmingDelete(false)}
      />
    </>
  );
}
