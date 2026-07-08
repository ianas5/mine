import { useMemo, useState, type ReactNode } from 'react';
import { SectionList, Text } from 'react-native';

import { useTheme } from '@/core/theme';
import { Input, ListRow, Sheet } from '@/core/ui';
import { LOAD_TYPE_LABELS } from '@/domain/fitness';
import type { Exercise } from '@/domain/models';

import { useExercises } from '../hooks/useExercises';
import { groupExercises } from '../logic/groupExercises';

interface ExercisePickerSheetProps {
  readonly visible: boolean;
  readonly onClose: () => void;
  readonly onPick: (exercise: Exercise) => void;
}

/** Fast exercise picker for the active workout — search + grouped, one tap to add. */
export function ExercisePickerSheet(props: ExercisePickerSheetProps): ReactNode {
  const theme = useTheme();
  const [query, setQuery] = useState('');
  const { exercises } = useExercises(false);
  const sections = useMemo(() => groupExercises(exercises ?? [], query), [exercises, query]);

  return (
    <Sheet visible={props.visible} onClose={props.onClose} title="Add Exercise">
      <Input
        value={query}
        onChangeText={setQuery}
        placeholder="Search exercises…"
        accessibilityLabel="Search exercises to add"
      />
      <SectionList
        style={{ maxHeight: 420, marginTop: theme.space.md }}
        sections={sections}
        keyExtractor={(item) => item.id}
        keyboardShouldPersistTaps="handled"
        renderSectionHeader={({ section }) => (
          <Text
            style={{
              ...theme.type.micro,
              color: theme.color.textSecondary,
              textTransform: 'uppercase',
              backgroundColor: theme.color.surface,
              paddingTop: theme.space.md,
              paddingBottom: theme.space.xs,
            }}
          >
            {section.title}
          </Text>
        )}
        renderItem={({ item }) => (
          <ListRow
            title={item.name}
            subtitle={LOAD_TYPE_LABELS[item.loadType]}
            onPress={() => {
              props.onPick(item);
              props.onClose();
            }}
          />
        )}
      />
    </Sheet>
  );
}
