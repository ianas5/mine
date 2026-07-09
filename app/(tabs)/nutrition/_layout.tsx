import { Stack } from 'expo-router';
import type { ReactNode } from 'react';

export default function NutritionStackLayout(): ReactNode {
  return <Stack screenOptions={{ headerShown: false }} />;
}
