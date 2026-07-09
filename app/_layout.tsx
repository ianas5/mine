import {
  Inter_400Regular,
  Inter_500Medium,
  Inter_600SemiBold,
  Inter_700Bold,
  useFonts,
} from '@expo-google-fonts/inter';
import { Stack } from 'expo-router';
import * as SplashScreen from 'expo-splash-screen';
import { StatusBar } from 'expo-status-bar';
import { useEffect, type ReactNode } from 'react';

import { DbGate } from '@/core/db';
import { ThemeProvider } from '@/core/theme';
import { ToastHost } from '@/core/ui';
import { PhotoSweeper } from '@/features/measurements/components/PhotoSweeper';
import { SessionKeeper } from '@/features/workouts/components/SessionKeeper';
// Composition root: data-layer artifacts are injected here so core never imports data.
import { setArchiveStore } from '@/data/backup';
import { expoArchiveStore } from '@/data/backup/expoArchiveStore';
import { expoPhotoStore } from '@/data/photos/expoPhotoStore';
import { setPhotoStore } from '@/data/photos/photoStore';
import { seedDatabase } from '@/data/seed/seedDatabase';
import migrations from '@/data/schema/migrations/migrations';

void SplashScreen.preventAutoHideAsync();

// Wire the real filesystem-backed stores once (like the SQLite handle).
setPhotoStore(expoPhotoStore);
setArchiveStore(expoArchiveStore);

export default function RootLayout(): ReactNode {
  const [fontsLoaded] = useFonts({
    Inter_400Regular,
    Inter_500Medium,
    Inter_600SemiBold,
    Inter_700Bold,
  });

  useEffect(() => {
    if (fontsLoaded) {
      void SplashScreen.hideAsync();
    }
  }, [fontsLoaded]);

  if (!fontsLoaded) {
    return null;
  }

  return (
    <ThemeProvider>
      <DbGate migrations={migrations} afterMigrate={seedDatabase}>
        {/* Crash-safety engine: recovery on launch + live draft checkpointing (§7.1). */}
        <SessionKeeper />
        {/* Reconcile photo files vs metadata on launch (DATABASE §3.6 orphan sweep). */}
        <PhotoSweeper />
        <StatusBar style="auto" />
        <Stack screenOptions={{ headerShown: false }}>
          <Stack.Screen name="(tabs)" />
          <Stack.Screen name="active-workout" options={{ animation: 'slide_from_bottom' }} />
        </Stack>
        <ToastHost />
      </DbGate>
    </ThemeProvider>
  );
}
