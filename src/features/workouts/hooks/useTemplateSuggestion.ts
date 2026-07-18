import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { isoWeekday, todayIso } from '@/core/utils';
import { suggestTemplate } from '@/domain/fitness';
import type { Template, Workout } from '@/domain/models';
import { programRepository } from '@/data/repositories/programRepository';
import { workoutRepository } from '@/data/repositories/workoutRepository';

export type ResolvedSuggestion =
  | { readonly kind: 'template'; readonly template: Template }
  | { readonly kind: 'repeatLast'; readonly workout: Workout }
  | { readonly kind: 'none' };

const LOOKBACK_DAYS = 56; // 8 weeks (UI_UX §5.2)

function daysAgoIso(days: number): string {
  const ms = new Date(`${todayIso()}T00:00:00`).getTime() - days * 86_400_000;
  return todayIso(new Date(ms));
}

/**
 * The smart-default start suggestion for the Workouts home (UI_UX §5.2): active
 * program's weekday template → most-frequent template on this weekday → Repeat
 * Last → nothing. Reactive to program and workout writes. `undefined` while loading.
 */
export function useTemplateSuggestion(): ResolvedSuggestion | undefined {
  const version = useTableVersion('programs', 'workouts');
  const [suggestion, setSuggestion] = useState<ResolvedSuggestion | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void (async () => {
      const weekday = isoWeekday(todayIso());

      const activeProgram = await programRepository.getActiveProgram();
      // Templates come ordered by position, so find() yields the first scheduled today.
      const scheduled = activeProgram?.templates.find((t) => t.weekdays.includes(weekday)) ?? null;

      const recent = await programRepository.getRecentTemplateUses(daysAgoIso(LOOKBACK_DAYS));
      const lastWorkouts = await workoutRepository.listRecent(1);
      const lastWorkout = lastWorkouts[0] ?? null;

      const decision = suggestTemplate(
        weekday,
        scheduled ? scheduled.id : null,
        recent,
        lastWorkout !== null,
      );

      let resolved: ResolvedSuggestion = { kind: 'none' };
      if (decision.kind === 'template') {
        const template =
          scheduled?.id === decision.templateId
            ? scheduled
            : await programRepository.getTemplate(decision.templateId);
        if (template) resolved = { kind: 'template', template };
        else if (lastWorkout) resolved = { kind: 'repeatLast', workout: lastWorkout };
      } else if (decision.kind === 'repeatLast' && lastWorkout) {
        resolved = { kind: 'repeatLast', workout: lastWorkout };
      }

      if (live) setSuggestion(resolved);
    })();
    return () => {
      live = false;
    };
  }, [version]);

  return suggestion;
}
