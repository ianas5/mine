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
// Composition root: data-layer artifacts are injected here so core never imports data.
import { seedDatabase } from '@/data/seed/seedDatabase';
import migrations from '@/data/schema/migrations/migrations';

void SplashScreen.preventAutoHideAsync();

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
