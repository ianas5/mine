import { Tabs } from 'expo-router';
// expo-router does not re-export the default tab bar at its top level; this is the
// canonical component behind Tabs, imported so a custom `tabBar` can dock the
// session bar directly above it (content reflows above both — no floating overlay).
import { BottomTabBar } from 'expo-router/build/react-navigation/bottom-tabs';
import type { BottomTabBarProps } from 'expo-router/build/react-navigation/bottom-tabs';
import { ChartLine, Dumbbell, LayoutDashboard, Ruler, UtensilsCrossed } from 'lucide-react-native';
import type { ReactNode } from 'react';
import { View } from 'react-native';

import { useTheme } from '@/core/theme';
import { SessionBar } from '@/features/workouts/components/SessionBar';

export default function TabsLayout(): ReactNode {
  const theme = useTheme();
  return (
    <Tabs
      tabBar={(props: BottomTabBarProps) => (
        <View>
          <SessionBar />
          <BottomTabBar {...props} />
        </View>
      )}
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
