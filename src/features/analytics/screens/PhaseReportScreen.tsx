import { useLocalSearchParams, useRouter } from 'expo-router';
import { ArrowLeft, Pencil, Trash2 } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Dialog, EmptyState, IconButton, Screen, Skeleton } from '@/core/ui';
import { phaseRepository } from '@/data/repositories/phaseRepository';

import { PhaseFormSheet } from '../components/PhaseFormSheet';
import { PhaseReportView } from '../components/PhaseReportView';
import { usePhaseReport } from '../hooks/usePhases';

/** The full Phase Report for one phase (§5.4), with edit + delete. */
export function PhaseReportScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const { id, celebrate } = useLocalSearchParams<{ id: string; celebrate?: string }>();
  const report = usePhaseReport(id ?? null);

  const [editing, setEditing] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);

  const del = async (): Promise<void> => {
    setConfirmDelete(false);
    if (id) await phaseRepository.deletePhase(id);
    router.back();
  };

  return (
    <Screen scroll>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.sm,
          marginTop: theme.space.sm,
          marginBottom: theme.space.xl,
        }}
      >
        <IconButton
          icon={<ArrowLeft color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
          onPress={() => router.back()}
          accessibilityLabel="Back"
        />
        <View style={{ flex: 1 }} />
        {report ? (
          <>
            <IconButton
              icon={<Pencil color={theme.color.textPrimary} size={20} strokeWidth={1.75} />}
              onPress={() => setEditing(true)}
              accessibilityLabel="Edit phase"
            />
            <IconButton
              icon={<Trash2 color={theme.color.danger} size={20} strokeWidth={1.75} />}
              onPress={() => setConfirmDelete(true)}
              accessibilityLabel="Delete phase"
            />
          </>
        ) : null}
      </View>

      {report === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={80} />
          <Skeleton height={160} />
          <Skeleton height={160} />
        </View>
      ) : report === null ? (
        <EmptyState title="This phase no longer exists." />
      ) : (
        <PhaseReportView report={report} celebrate={celebrate === '1'} />
      )}

      {report ? (
        <PhaseFormSheet
          visible={editing}
          editing={report.phase}
          onClose={() => setEditing(false)}
        />
      ) : null}

      <Dialog
        visible={confirmDelete}
        title="Delete this phase?"
        message="The phase is removed. Your workouts, meals and measurements are untouched."
        confirmLabel="Delete"
        cancelLabel="Cancel"
        onConfirm={() => void del()}
        onCancel={() => setConfirmDelete(false)}
      />
    </Screen>
  );
}
