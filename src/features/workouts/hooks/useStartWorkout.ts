import { useRouter } from 'expo-router';

import type { Template, Workout } from '@/domain/models';

import { prepareRepeatLast, prepareTemplateStart } from '../logic/startPreparation';
import { useSessionActions } from '../stores/useSessionStore';

export interface StartWorkout {
  readonly startEmpty: () => void;
  readonly startFromTemplate: (template: Template) => Promise<void>;
  readonly startRepeatLast: (workout: Workout) => void;
}

/**
 * The start-a-session entry points (Phase 8), each ≤ 1 tap from where it is shown:
 * begin the session and open the Active Workout screen. Template starts record the
 * template as provenance; nothing here ever writes back to a template or a workout.
 */
export function useStartWorkout(): StartWorkout {
  const router = useRouter();
  const actions = useSessionActions();
  const go = (): void => router.push('/active-workout');

  return {
    startEmpty: () => {
      actions.start(Date.now(), 'Workout');
      go();
    },
    startFromTemplate: async (template) => {
      const { name, prepared } = await prepareTemplateStart(template);
      actions.begin(Date.now(), name, prepared, template.id);
      go();
    },
    startRepeatLast: (workout) => {
      const { name, prepared } = prepareRepeatLast(workout);
      actions.begin(Date.now(), name, prepared, null);
      go();
    },
  };
}
