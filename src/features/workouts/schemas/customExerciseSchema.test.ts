import { customExerciseSchema } from './customExerciseSchema';

describe('customExerciseSchema', () => {
  const valid = {
    name: 'Cable Y-Raise',
    primaryMuscleGroup: 'shoulders',
    loadType: 'external',
    defaultUnilateral: false,
  };

  it('accepts a valid custom exercise', () => {
    expect(customExerciseSchema.safeParse(valid).success).toBe(true);
  });

  it('trims and rejects an empty name', () => {
    const result = customExerciseSchema.safeParse({ ...valid, name: '   ' });
    expect(result.success).toBe(false);
  });

  it('rejects an unknown muscle group', () => {
    expect(customExerciseSchema.safeParse({ ...valid, primaryMuscleGroup: 'delts' }).success).toBe(
      false,
    );
  });

  it('rejects an unknown load type', () => {
    expect(customExerciseSchema.safeParse({ ...valid, loadType: 'machine' }).success).toBe(false);
  });
});
