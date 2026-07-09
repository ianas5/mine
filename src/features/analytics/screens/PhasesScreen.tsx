import { useRouter, type Href } from 'expo-router';
import { ArrowLeft, Plus } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, Dialog, EmptyState, IconButton, ListRow, Screen, Skeleton } from '@/core/ui';
import { addDaysIso, todayIso } from '@/core/utils';
import { phaseRepository } from '@/data/repositories/phaseRepository';
import { PHASE_TYPE_LABELS, type Phase } from '@/domain/models';

import { PhaseFormSheet } from '../components/PhaseFormSheet';
import { usePhases } from '../hooks/usePhases';

const reportHref = (id: string): Href => `/analytics/phase/${id}` as Href;

/** Phase management (ANALYTICS §5.4): the ongoing block up top with end/edit, then the
 * history. Declaring over an overlapping range is caught and guided by the form. */
export function PhasesScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const phases = usePhases();

  const [formVisible, setFormVisible] = useState(false);
  const [editing, setEditing] = useState<Phase | null>(null);
  const [endTarget, setEndTarget] = useState<Phase | null>(null);
  const [completed, setCompleted] = useState<Phase | null>(null);

  const ongoing = phases?.find((p) => p.endDate === null) ?? null;
  const past = phases?.filter((p) => p.endDate !== null) ?? [];

  const declare = (): void => {
    setEditing(null);
    setFormVisible(true);
  };

  const confirmEnd = async (): Promise<void> => {
    const phase = endTarget;
    if (!phase) return;
    const yesterday = addDaysIso(todayIso(), -1);
    const endDate = yesterday >= phase.startDate ? yesterday : phase.startDate;
    setEndTarget(null);
    await phaseRepository.endPhase(phase.id, endDate);
    setCompleted(phase); // the phase-complete moment
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
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary, flex: 1 }}>Phases</Text>
        <IconButton
          icon={<Plus color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
          onPress={declare}
          accessibilityLabel="Declare a phase"
        />
      </View>

      {phases === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={120} />
          <Skeleton height={64} />
        </View>
      ) : (
        <View style={{ gap: theme.space.xl }}>
          <View style={{ gap: theme.space.md }}>
            <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
              CURRENT PHASE
            </Text>
            {ongoing ? (
              <Card style={{ gap: theme.space.md }}>
                <View>
                  <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
                    {PHASE_TYPE_LABELS[ongoing.type].toUpperCase()} · since {ongoing.startDate}
                  </Text>
                  <Text style={{ ...theme.type.heading, color: theme.color.textPrimary }}>
                    {ongoing.name}
                  </Text>
                </View>
                <View style={{ flexDirection: 'row', gap: theme.space.md }}>
                  <View style={{ flex: 1 }}>
                    <Button
                      label="View report"
                      variant="secondary"
                      size="md"
                      onPress={() => router.push(reportHref(ongoing.id))}
                    />
                  </View>
                  <View style={{ flex: 1 }}>
                    <Button
                      label="End phase"
                      variant="ghost"
                      size="md"
                      onPress={() => setEndTarget(ongoing)}
                    />
                  </View>
                </View>
              </Card>
            ) : (
              <Card style={{ gap: theme.space.md }}>
                <Text style={{ ...theme.type.body, color: theme.color.textSecondary }}>
                  No phase is running. Declare a cut, bulk, recomp or maintenance block to judge it
                  against its intent.
                </Text>
                <Button label="Declare a phase" onPress={declare} />
              </Card>
            )}
          </View>

          <View style={{ gap: theme.space.md }}>
            <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>HISTORY</Text>
            {past.length === 0 ? (
              <EmptyState title="Completed phases will appear here." />
            ) : (
              <Card>
                {past.map((p) => (
                  <ListRow
                    key={p.id}
                    title={p.name}
                    subtitle={`${PHASE_TYPE_LABELS[p.type]} · ${p.startDate} → ${p.endDate}`}
                    chevron
                    onPress={() => router.push(reportHref(p.id))}
                  />
                ))}
              </Card>
            )}
          </View>
        </View>
      )}

      <PhaseFormSheet
        visible={formVisible}
        editing={editing}
        onClose={() => setFormVisible(false)}
      />

      <Dialog
        visible={endTarget !== null}
        title="End this phase?"
        message={
          endTarget
            ? `${endTarget.name} will be marked ended, and you can review it as a completed block.`
            : ''
        }
        confirmLabel="End phase"
        cancelLabel="Keep going"
        onConfirm={() => void confirmEnd()}
        onCancel={() => setEndTarget(null)}
      />

      <Dialog
        visible={completed !== null}
        title={completed ? `${completed.name} complete` : ''}
        message="Nice work finishing the block. Your Phase Report is ready."
        confirmLabel="View report"
        cancelLabel="Done"
        onConfirm={() => {
          const id = completed?.id;
          setCompleted(null);
          if (id) router.push(`/analytics/phase/${id}?celebrate=1` as Href);
        }}
        onCancel={() => setCompleted(null)}
      />
    </Screen>
  );
}
