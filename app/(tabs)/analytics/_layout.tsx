import { Stack } from 'expo-router';
import type { ReactNode } from 'react';

export default function AnalyticsStackLayout(): ReactNode {
  return <Stack screenOptions={{ headerShown: false }} />;
}
