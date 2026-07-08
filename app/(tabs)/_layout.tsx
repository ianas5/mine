import { Tabs } from 'expo-router';
import { ChartLine, Dumbbell, LayoutDashboard, Ruler, UtensilsCrossed } from 'lucide-react-native';
import type { ReactNode } from 'react';

import { useTheme } from '@/core/theme';

export default function TabsLayout(): ReactNode {
  const theme = useTheme();
  return (
    <Tabs
      screenOptions={{
        headerShown: false,
        tabBarActiveTintColor: theme.color.accent,
        tabBarInactiveTintColor: theme.color.textTertiary,
        tabBarStyle: {
          backgroundColor: theme.color.surface,
          borderTopColor: theme.color.border,
        },
        tabBarLabelStyle: {
          fontFamily: theme.type.micro.fontFamily,
          fontSize: theme.type.micro.fontSize,
        },
      }}
    >
      <Tabs.Screen
        name="index"
        options={{
          title: 'Dashboard',
          tabBarIcon: ({ color }) => <LayoutDashboard color={color} size={24} strokeWidth={1.75} />,
        }}
      />
      <Tabs.Screen
        name="workouts"
        options={{
          title: 'Workouts',
          tabBarIcon: ({ color }) => <Dumbbell color={color} size={24} strokeWidth={1.75} />,
        }}
      />
      <Tabs.Screen
        name="nutrition"
        options={{
          title: 'Nutrition',
          tabBarIcon: ({ color }) => <UtensilsCrossed color={color} size={24} strokeWidth={1.75} />,
        }}
      />
      <Tabs.Screen
        name="measurements"
        options={{
          title: 'Measurements',
          tabBarIcon: ({ color }) => <Ruler color={color} size={24} strokeWidth={1.75} />,
        }}
      />
      <Tabs.Screen
        name="analytics"
        options={{
          title: 'Analytics',
          tabBarIcon: ({ color }) => <ChartLine color={color} size={24} strokeWidth={1.75} />,
        }}
      />
    </Tabs>
  );
}
