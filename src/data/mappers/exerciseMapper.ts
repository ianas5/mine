import type { LoadType, MuscleGroup } from '@/domain/fitness';
import type { Exercise } from '@/domain/models';

import type { ExerciseRow } from '../schema/tables';

export function rowToExercise(row: ExerciseRow): Exercise {
  return {
    id: row.id,
    name: row.name,
    primaryMuscleGroup: row.primaryMuscleGroup as MuscleGroup,
    secondaryMuscleGroups: parseGroups(row.secondaryMuscleGroups),
    loadType: row.loadType as LoadType,
    defaultUnilateral: row.defaultUnilateral === 1,
    isCustom: row.isCustom === 1,
    isArchived: row.isArchived === 1,
    notes: row.notes,
  };
}

function parseGroups(json: string): readonly MuscleGroup[] {
  try {
    const parsed: unknown = JSON.parse(json);
    return Array.isArray(parsed) ? (parsed as MuscleGroup[]) : [];
  } catch {
    return [];
  }
}
