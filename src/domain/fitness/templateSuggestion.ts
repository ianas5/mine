/** One past workout that was started from a template, for the weekday-mode fallback. */
export interface RecentTemplateUse {
  /** 0 = Monday … 6 = Sunday. */
  readonly weekday: number;
  readonly templateId: string;
}

export type TemplateSuggestion =
  | { readonly kind: 'template'; readonly templateId: string }
  | { readonly kind: 'repeatLast' }
  | { readonly kind: 'none' };

/**
 * The smart default for "what should I train?" (UI_UX §5.2), a pure function over
 * history — personalization without AI. The chain is: the active program's template
 * scheduled for today's weekday → else the most-frequent template used on this
 * weekday over recent history → else Repeat Last → else nothing. All fallbacks; the
 * user is never locked in.
 *
 * @param scheduledTemplateId the active program's template for today, resolved by
 *   the caller (null when there is no active program or none is scheduled today)
 * @param recent template-started workouts in the lookback window, newest first,
 *   already filtered to still-selectable templates
 */
export function suggestTemplate(
  weekday: number,
  scheduledTemplateId: string | null,
  recent: readonly RecentTemplateUse[],
  hasLastWorkout: boolean,
): TemplateSuggestion {
  if (scheduledTemplateId !== null) {
    return { kind: 'template', templateId: scheduledTemplateId };
  }

  const onWeekday = recent.filter((use) => use.weekday === weekday);
  if (onWeekday.length > 0) {
    const counts = new Map<string, number>();
    for (const use of onWeekday) counts.set(use.templateId, (counts.get(use.templateId) ?? 0) + 1);
    // Most frequent; ties resolve to the most recent (recent is newest-first, and
    // the first-seen winner is only replaced by a strictly higher count).
    let bestId = onWeekday[0]!.templateId;
    let bestCount = 0;
    for (const use of onWeekday) {
      const count = counts.get(use.templateId)!;
      if (count > bestCount) {
        bestCount = count;
        bestId = use.templateId;
      }
    }
    return { kind: 'template', templateId: bestId };
  }

  return hasLastWorkout ? { kind: 'repeatLast' } : { kind: 'none' };
}
