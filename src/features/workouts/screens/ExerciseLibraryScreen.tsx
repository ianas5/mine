import { useMemo, useState, type ReactNode } from 'react';
import { SectionList, Text, View } from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';

import { useTheme } from '@/core/theme';
import { Button, Chip, EmptyState, Input, ListRow, Skeleton } from '@/core/ui';
import { LOAD_TYPE_LABELS } from '@/domain/fitness';
import type { Exercise } from '@/domain/models';

import { CustomExerciseSheet } from '../components/CustomExerciseSheet';
import { ExerciseActionsSheet } from '../components/ExerciseActionsSheet';
import { useExercises } from '../hooks/useExercises';
import { groupExercises } from '../logic/groupExercises';

/** Exercise Library — browse/search the catalog, add custom, archive (Phase 3). */
export function ExerciseLibraryScreen(): ReactNode {
  const theme = useTheme();
  const insets = useSafeAreaInsets();
  const [query, setQuery] = useState('');
  const [showArchived, setShowArchived] = useState(false);
  const [customOpen, setCustomOpen] = useState(false);
  const [selected, setSelected] = useState<Exercise | null>(null);

  const { exercises } = useExercises(showArchived);
  const sections = useMemo(() => groupExercises(exercises ?? [], query), [exercises, query]);

  return (
    <View style={{ flex: 1, backgroundColor: theme.color.bg, paddingTop: insets.top }}>
      <View style={{ paddingHorizontal: theme.space.lg, gap: theme.space.md }}>
        <View
          style={{
            flexDirection: 'row',
            alignItems: 'center',
            justifyContent: 'space-between',
            marginTop: theme.space.sm,
          }}
        >
          <Text style={{ ...theme.type.title, color: theme.color.textPrimary }}>Exercises</Text>
          <Button label="New" size="md" onPress={() => setCustomOpen(true)} variant="secondary" />
        </View>
        <Input
          value={query}
          onChangeText={setQuery}
          placeholder="Search exercises…"
          accessibilityLabel="Search exercises"
        />
        <View style={{ flexDirection: 'row', gap: theme.space.sm }}>
          <Chip label="Active" selected={!showArchived} onPress={() => setShowArchived(false)} />
          <Chip label="Archived" selected={showArchived} onPress={() => setShowArchived(true)} />
        </View>
      </View>

      {exercises === null ? (
        <View style={{ padding: theme.space.lg, gap: theme.space.sm }}>
          <Skeleton height={44} />
          <Skeleton height={44} />
          <Skeleton height={44} />
        </View>
      ) : (
        <SectionList
          sections={sections}
          keyExtractor={(item) => item.id}
          contentContainerStyle={{
            paddingHorizontal: theme.space.lg,
            paddingBottom: insets.bottom + theme.space.xxl,
          }}
          keyboardShouldPersistTaps="handled"
          renderSectionHeader={({ section }) => (
            <Text
              style={{
                ...theme.type.micro,
                color: theme.color.textSecondary,
                textTransform: 'uppercase',
                backgroundColor: theme.color.bg,
                paddingTop: theme.space.lg,
                paddingBottom: theme.space.xs,
              }}
            >
              {section.title}
            </Text>
          )}
          renderItem={({ item }) => (
            <ListRow
              title={item.name}
              subtitle={
                LOAD_TYPE_LABELS[item.loadType] + (item.defaultUnilateral ? ' · unilateral' : '')
              }
              chevron
              onPress={() => setSelected(item)}
            />
          )}
          ListEmptyComponent={
            <EmptyState
              title={
                showArchived
                  ? 'No archived exercises'
                  : query
                    ? `No exercises match "${query}"`
                    : 'No exercises yet'
              }
              cta={
                showArchived
                  ? undefined
                  : { label: 'New exercise', onPress: () => setCustomOpen(true) }
              }
            />
          }
        />
      )}

      <CustomExerciseSheet visible={customOpen} onClose={() => setCustomOpen(false)} />
      <ExerciseActionsSheet exercise={selected} onClose={() => setSelected(null)} />
    </View>
  );
}
