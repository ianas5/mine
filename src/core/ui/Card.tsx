import type { ReactNode } from 'react';
import { View, type ViewStyle } from 'react-native';

import { useTheme } from '@/core/theme';

type CardVariant = 'default' | 'raised' | 'accentEdge';

interface CardProps {
  readonly children: ReactNode;
  readonly variant?: CardVariant;
  readonly style?: ViewStyle;
}

/** Surface card — variants per DESIGN_SYSTEM §6 (accentEdge is InsightCard's). */
export function Card(props: CardProps): ReactNode {
  const theme = useTheme();
  const variant = props.variant ?? 'default';
  const style: ViewStyle = {
    backgroundColor: variant === 'raised' ? theme.color.surfaceRaised : theme.color.surface,
    borderRadius: theme.radius.lg,
    padding: theme.space.lg,
    borderWidth: 1,
    borderColor: theme.color.border,
    ...(variant === 'accentEdge' && {
      borderLeftWidth: 3,
      borderLeftColor: theme.color.accent,
    }),
    ...(variant === 'raised' &&
      theme.mode === 'light' && {
        shadowColor: '#000000',
        shadowOpacity: 0.08,
        shadowRadius: 8,
        shadowOffset: { width: 0, height: 2 },
        elevation: 2,
      }),
    ...props.style,
  };
  return <View style={style}>{props.children}</View>;
}
