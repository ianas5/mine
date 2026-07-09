import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { formatKg, formatRelativeDate } from '@/core/utils';
import type { ExercisePreview, LoadType } from '@/domain/fitness';

import { formatSet } from '../logic/formatSet';

interface ExerciseHistoryPanelProps {
  readonly preview: ExercisePreview | null;
  readonly loadType: LoadType;
}

/** Compact Last / Best / Best e1RM at the point of logging (UI_UX §4.1). */
export function ExerciseHistoryPanel(props: ExerciseHistoryPanelProps): ReactNode {
  const theme = useTheme();
  const { preview, loadType } = props;

  const labelStyle = { ...theme.type.micro, color: theme.color.textTertiary } as const;
  const valueStyle = { ...theme.type.caption, color: theme.color.textSecondary } as const;

  if (preview === null) {
    return null;
  }
  if (preview.last === null) {
    return (
      <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>
        First time — set your baseline
      </Text>
    );
  }

  const row = (label: string, value: string): ReactNode => (
    <View style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.sm }}>
      <Text style={[labelStyle, { width: 64 }]}>{label}</Text>
      <Text style={valueStyle}>{value}</Text>
    </View>
  );

  return (
    <View
      style={{
        gap: theme.space.xs,
        paddingVertical: theme.space.sm,
        borderBottomWidth: 1,
        borderBottomColor: theme.color.border,
        marginBottom: theme.space.xs,
      }}
    >
      {row(
        `Last · ${formatRelativeDate(preview.last.date)}`.toUpperCase(),
        preview.last.sets.map((s) => formatSet(s.weightKg, s.reps, loadType)).join(' · '),
      )}
      {preview.bestWeightSet && preview.bestWeightSet.weightKg > 0
        ? row(
            'BEST',
            formatSet(preview.bestWeightSet.weightKg, preview.bestWeightSet.reps, loadType),
          )
        : null}
      {preview.bestE1rmKg !== null ? row('BEST E1RM', `${formatKg(preview.bestE1rmKg)} kg`) : null}
    </View>
  );
}
