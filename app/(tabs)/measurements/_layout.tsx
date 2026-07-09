import { Stack } from 'expo-router';
import type { ReactNode } from 'react';

export default function MeasurementsStackLayout(): ReactNode {
  return <Stack screenOptions={{ headerShown: false }} />;
}
