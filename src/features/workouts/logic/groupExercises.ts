import { MUSCLE_GROUPS, MUSCLE_GROUP_LABELS, type MuscleGroup } from '@/domain/fitness';
import type { Exercise } from '@/domain/models';

export interface ExerciseSection {
  readonly group: MuscleGroup;
  readonly title: string;
  readonly data: readonly Exercise[];
}

/**
 * Groups exercises by primary muscle (canonical order) and applies a
 * case-insensitive name search. Empty groups are dropped. Pure (ARCHITECTURE §9.1).
 */
export function groupExercises(list: readonly Exercise[], query: string): ExerciseSection[] {
  const q = query.trim().toLowerCase();
  const filtered = q ? list.filter((e) => e.name.toLowerCase().includes(q)) : list;
  return MUSCLE_GROUPS.map((group) => ({
    group,
    title: MUSCLE_GROUP_LABELS[group],
    data: filtered.filter((e) => e.primaryMuscleGroup === group),
  })).filter((section) => section.data.length > 0);
}
