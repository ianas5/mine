import { useRouter } from 'expo-router';
import { ArrowLeft, ChevronRight } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, EmptyState, IconButton, Input, Screen, Sheet, Skeleton } from '@/core/ui';
import { programRepository } from '@/data/repositories/programRepository';

import { usePrograms } from '../hooks/usePrograms';

/** Programs list — create a program or open one to plan its sessions (Phase 8). */
export function ProgramsScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const programs = usePrograms();
  const [creating, setCreating] = useState(false);
  const [newName, setNewName] = useState('');

  const create = async (): Promise<void> => {
    const id = await programRepository.createProgram({ name: newName, notes: null });
    setCreating(false);
    setNewName('');
    router.push(`/workouts/programs/${id}`);
  };

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
          Programs
        </Text>
        <Button label="New" size="md" variant="secondary" onPress={() => setCreating(true)} />
      </View>

      {programs === undefined ? (
        <View style={{ gap: theme.space.sm }}>
          <Skeleton height={64} />
          <Skeleton height={64} />
        </View>
      ) : programs.length === 0 ? (
        <EmptyState
          title="No programs yet"
          cta={{ label: 'New program', onPress: () => setCreating(true) }}
        />
      ) : (
        <View style={{ gap: theme.space.sm }}>
          {programs.map((program) => (
            <Pressable
              key={program.id}
              onPress={() => router.push(`/workouts/programs/${program.id}`)}
              accessibilityRole="button"
              accessibilityLabel={program.name}
              style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
            >
              <Card
                variant={program.isActive ? 'accentEdge' : 'default'}
                style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}
              >
                <View style={{ flex: 1 }}>
                  <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
                    {program.name}
                  </Text>
                  <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
                    {program.isActive ? 'Active · ' : ''}
                    {program.templates.length}{' '}
                    {program.templates.length === 1 ? 'session' : 'sessions'}
                  </Text>
                </View>
                <ChevronRight color={theme.color.textTertiary} size={20} />
              </Card>
            </Pressable>
          ))}
        </View>
      )}

      <Sheet visible={creating} onClose={() => setCreating(false)} title="New program">
        <View style={{ gap: theme.space.lg }}>
          <Input
            value={newName}
            onChangeText={setNewName}
            placeholder="Program name (e.g. Push Pull Legs)"
            accessibilityLabel="Program name"
            autoFocus
          />
          <Button label="Create program" onPress={() => void create()} />
        </View>
      </Sheet>
    </Screen>
  );
}
