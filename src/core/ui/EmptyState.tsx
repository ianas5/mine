import type { ReactNode } from 'react';
import { Text, View } from 'react-native';

import { useTheme } from '@/core/theme';

import { Button } from './Button';

interface EmptyStateProps {
  readonly icon?: ReactNode;
  readonly title: string;
  readonly cta?: { readonly label: string; readonly onPress: () => void } | undefined;
}

/** Quiet, directive empty state — one factual line, optional action (UI_UX §6). */
export function EmptyState(props: EmptyStateProps): ReactNode {
  const theme = useTheme();
  return (
    <View
      style={{
        alignItems: 'center',
        justifyContent: 'center',
        paddingVertical: theme.space.xxxl,
        gap: theme.space.md,
      }}
    >
      {props.icon ?? null}
      <Text
        style={{ ...theme.type.body, color: theme.color.textSecondary, textAlign: 'center' }}
        accessibilityRole="text"
      >
        {props.title}
      </Text>
      {props.cta ? (
        <Button variant="secondary" size="md" label={props.cta.label} onPress={props.cta.onPress} />
      ) : null}
    </View>
  );
}
