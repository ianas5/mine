import { useRouter, type Href } from 'expo-router';
import type { ReactNode } from 'react';
import { Pressable, Text } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card } from '@/core/ui';
import { PHASE_TYPE_LABELS } from '@/domain/models';

import { usePhaseReport } from '../hooks/usePhases';

/** The current-phase progress card for the Analytics home (§5.4). Shows the ongoing block's
 * progress-to-date, or a calm prompt to declare one. Taps through to phase management. */
export function CurrentPhaseCard(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const report = usePhaseReport(null);

  if (report === undefined) return null;

  const go = (): void => router.push('/analytics/phases' as Href);

  if (report === null) {
    return (
      <Pressable onPress={go} accessibilityRole="button" accessibilityLabel="Declare a phase">
        <Card style={{ gap: theme.space.xs }}>
          <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
            TRAINING PHASE
          </Text>
          <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
            No active phase
          </Text>
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            Declare a cut, bulk or recomp to judge a block against its intent.
          </Text>
        </Card>
      </Pressable>
    );
  }

  const { phase, training } = report;
  return (
    <Pressable onPress={go} accessibilityRole="button" accessibilityLabel={`Phase ${phase.name}`}>
      <Card style={{ gap: theme.space.xs }}>
        <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
          CURRENT PHASE · {PHASE_TYPE_LABELS[phase.type].toUpperCase()}
        </Text>
        <Text style={{ ...theme.type.heading, color: theme.color.textPrimary }}>{phase.name}</Text>
        <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
          {report.spanDays} days in · {training.workouts} workouts ·{' '}
          {report.nutrition.intent.message}
        </Text>
      </Card>
    </Pressable>
  );
}
