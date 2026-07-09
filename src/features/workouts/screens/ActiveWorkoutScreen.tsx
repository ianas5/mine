import { useRouter } from 'expo-router';
import { Plus, X } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, ScrollView, Text, TextInput, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { useTheme } from '@/core/theme';
import { Button, Dialog, EmptyState, showToast } from '@/core/ui';
import { formatElapsed } from '@/core/utils';
import { computeWorkoutStats } from '@/domain/fitness';
import type { Exercise } from '@/domain/models';
import { workoutRepository } from '@/data/repositories/workoutRepository';

import { ActiveExerciseCard } from '../components/ActiveExerciseCard';
import { ExercisePickerSheet } from '../components/ExercisePickerSheet';
import { WorkoutSummarySheet } from '../components/WorkoutSummarySheet';
import { useDefaultBodyweight } from '../hooks/useDefaultBodyweight';
import { useElapsed } from '../hooks/useElapsed';
import { sessionToStatExercises, sessionToWorkoutInput } from '../logic/sessionMapping';
import { useSessionActions, useSessionStore } from '../stores/useSessionStore';

/** The active workout — the heart of the app (Phase 4). Fast logging, clear progress. */
export function ActiveWorkoutScreen(): ReactNode {
  const theme = useTheme();
  const insets = useSafeAreaInsets();
  const router = useRouter();
  const actions = useSessionActions();

  const active = useSessionStore((s) => s.active);
  const name = useSessionStore((s) => s.name);
  const startedAt = useSessionStore((s) => s.startedAt);
  const exercises = useSessionStore((s) => s.exercises);

  const elapsedMs = useElapsed(startedAt);
  const bodyweightKg = useDefaultBodyweight();

  const [pickerOpen, setPickerOpen] = useState(false);
  const [summaryOpen, setSummaryOpen] = useState(false);
  const [confirmDiscard, setConfirmDiscard] = useState(false);
  const [saving, setSaving] = useState(false);

  if (!active) {
    return (
      <View style={{ flex: 1, backgroundColor: theme.color.bg, justifyContent: 'center' }}>
        <EmptyState
          title="No active workout"
          cta={{ label: 'Go back', onPress: () => router.back() }}
        />
      </View>
    );
  }

  const totalSets = exercises.reduce((n, ex) => n + ex.sets.length, 0);
  const doneSets = exercises.reduce((n, ex) => n + ex.sets.filter((s) => s.done).length, 0);
  const hasWorkingSets =
    computeWorkoutStats(sessionToStatExercises(exercises), bodyweightKg).workingSetCount > 0;

  const leave = (): void => {
    actions.discard();
    router.back();
  };

  const onSave = async (): Promise<void> => {
    setSaving(true);
    try {
      await workoutRepository.saveCompletedWorkout(
        sessionToWorkoutInput(useSessionStore.getState(), Date.now()),
      );
      setSummaryOpen(false);
      showToast('Workout saved', 'success');
      leave();
    } catch {
      showToast('Could not save workout');
      setSaving(false);
    }
  };

  // Prefill the added exercise with last time's sets so repeats need zero typing (Phase 5).
  const addExerciseWithHistory = async (exercise: Exercise): Promise<void> => {
    const preview = await workoutRepository.getExercisePreview(exercise.id, exercise.loadType);
    actions.addExercise(exercise, preview.last?.sets);
  };

  const requestDiscard = (): void => {
    setSummaryOpen(false);
    if (hasWorkingSets) {
      setConfirmDiscard(true);
    } else {
      leave();
    }
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
        <Pressable
          onPress={() => router.back()}
          accessibilityRole="button"
          accessibilityLabel="Minimize workout"
          hitSlop={theme.space.sm}
          style={{ width: 40, height: 40, alignItems: 'center', justifyContent: 'center' }}
        >
          <X color={theme.color.textPrimary} size={24} />
        </Pressable>
        <TextInput
          value={name}
          onChangeText={actions.setName}
          accessibilityLabel="Workout name"
          style={{ ...theme.type.heading, color: theme.color.textPrimary, flex: 1 }}
        />
        <Button label="Finish" size="md" onPress={() => setSummaryOpen(true)} />
      </View>

      <View
        style={{
          paddingHorizontal: theme.space.lg,
          paddingBottom: theme.space.md,
          flexDirection: 'row',
          alignItems: 'center',
          justifyContent: 'space-between',
        }}
      >
        <Text
          style={{
            ...theme.type.display,
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {formatElapsed(elapsedMs)}
        </Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
          {exercises.length} exercises · {doneSets}/{totalSets} sets done
        </Text>
      </View>

      <ScrollView
        contentContainerStyle={{
          paddingHorizontal: theme.space.lg,
          paddingBottom: insets.bottom + theme.space.xxxl,
          gap: theme.space.md,
        }}
        keyboardShouldPersistTaps="handled"
      >
        {exercises.length === 0 ? (
          <EmptyState
            title="Add your first exercise to start logging"
            cta={{ label: 'Add exercise', onPress: () => setPickerOpen(true) }}
          />
        ) : (
          <>
            {exercises.map((exercise) => (
              <ActiveExerciseCard key={exercise.localId} exercise={exercise} />
            ))}
            <Pressable
              onPress={() => setPickerOpen(true)}
              accessibilityRole="button"
              accessibilityLabel="Add exercise"
              style={({ pressed }) => ({
                flexDirection: 'row',
                alignItems: 'center',
                justifyContent: 'center',
                gap: theme.space.sm,
                paddingVertical: theme.space.lg,
                borderRadius: theme.radius.lg,
                backgroundColor: theme.color.surface,
                opacity: pressed ? 0.7 : 1,
              })}
            >
              <Plus color={theme.color.accent} size={20} />
              <Text style={{ ...theme.type.bodyStrong, color: theme.color.accent }}>
                Add exercise
              </Text>
            </Pressable>
          </>
        )}
      </ScrollView>

      <ExercisePickerSheet
        visible={pickerOpen}
        onClose={() => setPickerOpen(false)}
        onPick={(exercise: Exercise) => void addExerciseWithHistory(exercise)}
      />
      <WorkoutSummarySheet
        visible={summaryOpen}
        exercises={exercises}
        bodyweightKg={bodyweightKg}
        elapsedMs={elapsedMs}
        saving={saving}
        onSave={() => void onSave()}
        onDiscard={requestDiscard}
        onClose={() => setSummaryOpen(false)}
      />
      <Dialog
        visible={confirmDiscard}
        title="Discard workout?"
        message="This session and its logged sets will be lost."
        confirmLabel="Discard"
        onConfirm={() => {
          setConfirmDiscard(false);
          leave();
        }}
        onCancel={() => setConfirmDiscard(false)}
      />
    </View>
  );
}
