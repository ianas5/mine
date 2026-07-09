import { useRouter } from 'expo-router';
import { ArrowLeft, ChevronRight } from 'lucide-react-native';
import { useState, type ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Button, Card, EmptyState, IconButton, Input, Screen, Skeleton } from '@/core/ui';

import { useFoodPicks } from '../hooks/useFoodPicks';

/** Foods catalog — search, open one to edit, or create a new food (Phase 9). */
export function FoodsScreen(): ReactNode {
  const theme = useTheme();
  const router = useRouter();
  const picks = useFoodPicks();
  const [query, setQuery] = useState('');

  const filtered = (picks ?? []).filter((p) =>
    p.food.name.toLowerCase().includes(query.trim().toLowerCase()),
  );

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
        <Text style={{ ...theme.type.title, color: theme.color.textPrimary, flex: 1 }}>Foods</Text>
        <Button
          label="New"
          size="md"
          variant="secondary"
          onPress={() => router.push('/nutrition/foods/new')}
        />
      </View>

      <Input
        value={query}
        onChangeText={setQuery}
        placeholder="Search foods…"
        accessibilityLabel="Search foods"
      />

      <View style={{ marginTop: theme.space.md }}>
        {picks === undefined ? (
          <View style={{ gap: theme.space.sm }}>
            <Skeleton height={56} />
            <Skeleton height={56} />
          </View>
        ) : filtered.length === 0 ? (
          <EmptyState
            title={query ? `No foods match "${query}"` : 'No foods yet'}
            cta={{ label: 'New food', onPress: () => router.push('/nutrition/foods/new') }}
          />
        ) : (
          <View style={{ gap: theme.space.sm }}>
            {filtered.map((pick) => (
              <Pressable
                key={pick.food.id}
                onPress={() => router.push(`/nutrition/foods/${pick.food.id}`)}
                accessibilityRole="button"
                accessibilityLabel={pick.food.name}
                style={({ pressed }) => ({ opacity: pressed ? 0.7 : 1 })}
              >
                <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
                  <View style={{ flex: 1 }}>
                    <Text style={{ ...theme.type.body, color: theme.color.textPrimary }}>
                      {pick.food.name}
                      {pick.food.isQuickMeal ? '  ·  quick meal' : ''}
                    </Text>
                    <Text style={{ ...theme.type.caption, color: theme.color.textSecondary }}>
                      {pick.food.kcal} kcal · {pick.food.proteinG}P {pick.food.carbG}C{' '}
                      {pick.food.fatG}F per {pick.food.servingAmount} {pick.food.servingUnit}
                    </Text>
                  </View>
                  <ChevronRight color={theme.color.textTertiary} size={20} />
                </Card>
              </Pressable>
            ))}
          </View>
        )}
      </View>
    </Screen>
  );
}
