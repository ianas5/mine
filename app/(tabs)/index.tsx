import type { ReactNode } from 'react';

import {
  DashboardScreen,
  type ActiveSessionSummary,
} from '@/features/dashboard/screens/DashboardScreen';
import { useSessionStore } from '@/features/workouts/stores/useSessionStore';

/**
 * Composition root for the Dashboard tab: the active-session summary is read from the
 * workouts store here (app may import any feature) and passed as data, so the dashboard
 * feature never imports the workouts feature across the boundary (ARCHITECTURE §4).
 */
export default function DashboardRoute(): ReactNode {
  const active = useSessionStore((s) => s.active);
  const name = useSessionStore((s) => s.name);
  const startedAt = useSessionStore((s) => s.startedAt);
  const exercises = useSessionStore((s) => s.exercises);

  const summary: ActiveSessionSummary | null = active
    ? {
        name,
        startedAt,
        exerciseCount: exercises.length,
        completedSetCount: exercises.reduce(
          (n, ex) => n + ex.sets.filter((set) => set.done && !set.warmup).length,
          0,
        ),
      }
    : null;

  return <DashboardScreen activeSession={summary} />;
}
