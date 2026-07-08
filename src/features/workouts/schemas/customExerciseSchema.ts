import { z } from 'zod';

import { LOAD_TYPES, MUSCLE_GROUPS } from '@/domain/fitness';

/** Validation for the custom-exercise sheet (ARCHITECTURE §10; Phase 3). */
export const customExerciseSchema = z.object({
  name: z.string().trim().min(1, 'Name is required').max(60, 'Keep it under 60 characters'),
  primaryMuscleGroup: z.enum(MUSCLE_GROUPS),
  loadType: z.enum(LOAD_TYPES),
  defaultUnilateral: z.boolean(),
});

export type CustomExerciseInput = z.infer<typeof customExerciseSchema>;
