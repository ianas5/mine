import { useLocalSearchParams, useRouter } from 'expo-router';
import { ArrowLeft, ChevronRight, Trash2 } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, Dialog, EmptyState, IconButton, Input, Screen, Skeleton } from '@/core/ui';
import { weekdayLabel } from '@/core/utils';
import type { Program, Template } from '@/domain/models';
import { programRepository } from '@/data/repositories/programRepository';

import { useProgram } from '../hooks/usePrograms';

/** One program: set active, edit its session templates, delete (Phase 8). */
export function ProgramDetailScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const params = useLocalSearchParams<{ id: string }>();
  const program = useProgram(params.id ?? '');

  return (
    <Screen scroll>
      <View
        style={{
          flexDirection: 'row',
          alignItems: 'center',
          gap: theme.space.sm,
          marginTop: theme.space.sm,
          marginBottom: theme.space.lg,
        }}
      >
        <IconButton
          icon={<ArrowLeft color={theme.color.textPrimary} size={24} strokeWidth={1.75} />}
          onPress={() => router.back()}
          accessibilityLabel="Back"
        />
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary, flex: 1 }}>
          {program ? program.name : 'Program'}
        </Text>
      </View>

      {program === undefined ? (
        <View style={{ gap: theme.space.md }}>
          <Skeleton height={56} />
          <Skeleton height={96} />
        </View>
      ) : program === null ? (
        <EmptyState title="Program not found" cta={{ label: 'Go back', onPress: router.back }} />
      ) : (
        <LoadedProgram program={program} onDeleted={() => router.back()} />
      )}
    </Screen>
  );
}

function LoadedProgram(props: {
  readonly program: Program;
  readonly onDeleted: () => void;
}): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const { program } = props;
  // Seeded once on mount (the route's program id is fixed for this screen), so
  // reactive program refreshes never clobber what the user is typing.
  const [name, setName] = useState(program.name);
  const [confirmDelete, setConfirmDelete] = useState(false);

  const commitName = (): void => {
    if (name.trim() && name.trim() !== program.name) {
      void programRepository.updateProgram(program.id, { name });
    }
  };

  const addSession = async (): Promise<void> => {
    const id = await programRepository.createTemplate(program.id, {
      name: 'New session',
      weekday: null,
      notes: null,
      exercises: [],
    });
    router.push(`/workouts/templates/${id}`);
  };

  return (
    <View style={{ gap: theme.space.lg }}>
      <Input
        value={name}
        onChangeText={setName}
        onBlur={commitName}
        accessibilityLabel="Program name"
        placeholder="Program name"
      />

      {program.isActive ? (
        <Card
          variant="accentEdge"
          style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}
        >
          <View style={{ flex: 1 }}>
            <Text style={{ ...theme.type.bodyStrong, color: theme.color.textPrimary }}>
              Active program
            </Text>
            <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
              Its weekday sessions suggest your next workout.
            </Text>
          </View>
          <Button
            label="Deactivate"
            variant="secondary"
            size="md"
            onPress={() => void programRepository.clearActive()}
          />
        </Card>
      ) : (
        <Button
          label="Set as active program"
          onPress={() => void programRepository.setActive(program.id)}
        />
      )}

      <View style={{ gap: theme.space.sm }}>
        <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>SESSIONS</Text>
        {program.templates.length === 0 ? (
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
            No sessions yet. Add one to plan exercises and weekdays.
          </Text>
        ) : (
          program.templates.map((template) => (
            <TemplateRow
              key={template.id}
              template={template}
              onPress={() => router.push(`/workouts/templates/${template.id}`)}
            />
          ))
        )}
        <Button label="Add session" variant="secondary" onPress={() => void addSession()} />
      </View>

      <Pressable
        onPress={() => setConfirmDelete(true)}
        accessibilityRole="button"
        accessibilityLabel="Delete program"
        style={({ pressed }) => ({
          flexDirection: 'row',
          alignItems: 'center',
          justifyContent: 'center',
          gap: theme.space.xs,
          paddingVertical: theme.space.md,
          opacity: pressed ? 0.6 : 1,
        })}
      >
        <Trash2 color={theme.color.danger} size={18} />
        <Text style={{ ...theme.type.caption, color: theme.color.danger }}>Delete program</Text>
      </Pressable>

      <Dialog
        visible={confirmDelete}
        title="Delete program?"
        message="The program and its session templates are removed. Past workouts are kept — a template is a plan, not history."
        confirmLabel="Delete"
        onConfirm={() => {
          setConfirmDelete(false);
          void programRepository.deleteProgram(program.id).then(props.onDeleted);
        }}
        onCancel={() => setConfirmDelete(false)}
      />
    </View>
  );
}

function TemplateRow(props: {
  readonly template: Template;
  readonly onPress: () => void;
}): ReactNode {
  const theme = useTheme();
  const { template } = props;
  const meta = [
    template.weekday !== null ? weekdayLabel(template.weekday) : 'Any day',
    `${template.exercises.length} ${template.exercises.length === 1 ? 'exercise' : 'exercises'}`,
  ].join(' · ');

  return (
    <Pressable
      onPress={props.onPress}
      accessibilityRole="button"
      accessibilityLabel={template.name}
      style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
    >
      <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
        <View style={{ flex: 1 }}>
          <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
            {template.name}
          </Text>
          <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>{meta}</Text>
        </View>
        <ChevronRight color={theme.color.textTertiary} size={20} />
      </Card>
    </Pressable>
  );
}
