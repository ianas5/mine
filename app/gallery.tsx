import { Redirect } from 'expo-router';
import type { ReactNode } from 'react';

import { GalleryScreen } from '@/features/gallery/screens/GalleryScreen';

export default function GalleryRoute(): ReactNode {
  // Dev-builds only (IMPLEMENTATION_ROADMAP Phase 1); production builds redirect home.
  if (!__DEV__) {
    return <Redirect href="/" />;
  }
  return <GalleryScreen />;
}
