import { Droplet, Minus, Plus } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { Pressable, Text, View } from 'react-native';

import { useTheme } from '@/core/theme';
import { Card } from '@/core/ui';
import { nutritionRepository } from '@/data/repositories/nutritionRepository';

interface WaterCardProps {
  readonly date: string;
  /** Logged ml, or null when unlogged (0 ≠ absent, §2.4). */
  readonly waterMl: number | null;
  readonly targetMl: number | null;
  readonly cupMl: number;
}

/** Water logging (+/- a configurable cup). Preserves the 0-vs-unlogged distinction. */
export function WaterCard(props: WaterCardProps): ReactNode {
  const theme = useTheme();
  const { date, waterMl, targetMl, cupMl } = props;

  const hit = targetMl !== null && waterMl !== null && waterMl >= targetMl;
  const displayMl = waterMl ?? 0;

  const button = (icon: ReactNode, label: string, delta: number, disabled = false): ReactNode => (
    <Pressable
      onPress={() => void nutritionRepository.addWater(date, delta)}
      disabled={disabled}
      accessibilityRole="button"
      accessibilityLabel={label}
      style={({ pressed }) => ({
        width: 44,
        height: 44,
        borderRadius: theme.radius.full,
        alignItems: 'center',
        justifyContent: 'center',
        borderWidth: 1,
        borderColor: theme.color.border,
        opacity: disabled ? 0.4 : pressed ? 0.6 : 1,
      })}
    >
      {icon}
    </Pressable>
  );

  return (
    <Card style={{ flexDirection: 'row', alignItems: 'center', gap: theme.space.md }}>
      <Droplet
        color={hit ? theme.color.positive : theme.color.accent}
        size={22}
        strokeWidth={1.75}
      />
      <View style={{ flex: 1 }}>
        <Text
          style={{
            ...theme.type.bodyStrong,
            color: theme.color.textPrimary,
            fontVariant: ['tabular-nums'],
          }}
        >
          {waterMl === null ? 'Not logged' : `${displayMl.toLocaleString()} ml`}
          {targetMl !== null ? (
            <Text style={{ ...theme.type.caption, color: theme.color.textTertiary }}>
              {' '}
              / {targetMl.toLocaleString()} ml
            </Text>
          ) : null}
        </Text>
        <Text style={{ ...theme.type.micro, color: theme.color.textSecondary }}>
          WATER · {cupMl} ml per cup
        </Text>
      </View>
      {button(
        <Minus color={theme.color.textSecondary} size={20} />,
        'Remove a cup',
        -cupMl,
        displayMl <= 0,
      )}
      {button(<Plus color={theme.color.accent} size={20} />, 'Add a cup', cupMl)}
    </Card>
  );
}
