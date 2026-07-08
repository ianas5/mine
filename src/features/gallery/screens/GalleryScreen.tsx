import type { ReactNode } from 'react';
import { Text } from 'react-native';

import { useTheme } from '@/core/theme';
import { Screen } from '@/core/ui';

import { GalleryActions } from '../components/GalleryActions';
import { GalleryOverlays } from '../components/GalleryOverlays';
import { GalleryStructure } from '../components/GalleryStructure';

/** Dev-only showcase of every core/ui primitive in every state (Phase 1 gallery). */
export function GalleryScreen(): ReactNode {
  const theme = useTheme();
  return (
    <Screen scroll>
      <Text
        style={{
          ...theme.type.title,
          color: theme.color.textPrimary,
          marginTop: theme.space.lg,
          marginBottom: theme.space.xxl,
        }}
      >
        Design System Gallery
      </Text>
      <GalleryStructure />
      <GalleryActions />
      <GalleryOverlays />
    </Screen>
  );
}
