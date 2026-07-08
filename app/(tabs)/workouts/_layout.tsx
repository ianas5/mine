import { Stack } from 'expo-router';
import type { ReactNode } from 'react';

export default function WorkoutsStackLayout(): ReactNode {
  return <Stack screenOptions={{ headerShown: false }} />;
}
