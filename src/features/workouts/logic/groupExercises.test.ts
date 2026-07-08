import type { Exercise } from '@/domain/models';

import { groupExercises } from './groupExercises';

const make = (id: string, name: string, group: Exercise['primaryMuscleGroup']): Exercise => ({
  id,
  name,
  primaryMuscleGroup: group,
  secondaryMuscleGroups: [],
  loadType: 'external',
  defaultUnilateral: false,
  isCustom: false,
  isArchived: false,
  notes: null,
});

describe('groupExercises', () => {
  const list = [
    make('1', 'Bench Press', 'chest'),
    make('2', 'Cable Fly', 'chest'),
    make('3', 'Lat Pulldown', 'back'),
  ];

  it('groups by primary muscle in canonical order and drops empty groups', () => {
    const sections = groupExercises(list, '');

    expect(sections.map((s) => s.group)).toEqual(['chest', 'back']);
    expect(sections[0]?.data).toHaveLength(2);
  });

  it('filters case-insensitively by name', () => {
    const sections = groupExercises(list, 'cable');

    expect(sections).toHaveLength(1);
    expect(sections[0]?.data[0]?.name).toBe('Cable Fly');
  });

  it('returns nothing when the query matches no exercise', () => {
    expect(groupExercises(list, 'zzz')).toEqual([]);
  });
});
