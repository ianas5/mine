/** @jest-environment node */
import { setDbForTesting } from '@/core/db';

import { seedDatabase } from '../seed/seedDatabase';
import { createTestDb, type TestDb } from '../testing/createTestDb';
import { programRepository, type TemplateInput } from './programRepository';
import { workoutRepository } from './workoutRepository';

const BENCH = 'ex_seed_barbell-bench-press';
const SQUAT = 'ex_seed_back-squat';

const benchTemplate: TemplateInput = {
  name: 'Push A',
  weekdays: [0],
  notes: null,
  exercises: [
    {
      exerciseId: BENCH,
      targetSets: 3,
      targetRepMin: 8,
      targetRepMax: 10,
      targetRpe: 8,
      restSeconds: 120,
      notes: null,
    },
  ],
};

describe('programRepository (real SQLite)', () => {
  let testDb: TestDb;

  beforeEach(async () => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
    await seedDatabase();
  });

  afterEach(() => testDb.close());

  it('enforces a single active program', async () => {
    const a = await programRepository.createProgram({ name: 'A', notes: null });
    const b = await programRepository.createProgram({ name: 'B', notes: null });

    await programRepository.setActive(a);
    await programRepository.setActive(b);

    const active = await programRepository.getActiveProgram();
    expect(active?.id).toBe(b);
    const all = await programRepository.listPrograms();
    expect(all.filter((p) => p.isActive)).toHaveLength(1);
  });

  it('creates and resolves a template with its exercise targets', async () => {
    const programId = await programRepository.createProgram({ name: 'PPL', notes: null });
    const templateId = await programRepository.createTemplate(programId, benchTemplate);

    const template = await programRepository.getTemplate(templateId);
    expect(template?.weekdays).toEqual([0]);
    expect(template?.exercises).toHaveLength(1);
    expect(template?.exercises[0]).toMatchObject({
      exerciseId: BENCH,
      name: 'Barbell Bench Press',
      target: { sets: 3, repMin: 8, repMax: 10, rpe: 8, restSeconds: 120 },
    });
  });

  it('stores multiple weekdays and returns them sorted and de-duplicated', async () => {
    const templateId = await programRepository.createTemplate(null, {
      ...benchTemplate,
      weekdays: [4, 0, 4, 2], // unsorted, with a duplicate
    });
    const template = await programRepository.getTemplate(templateId);
    expect(template?.weekdays).toEqual([0, 2, 4]);
  });

  it('editing a template never rewrites a past workout (a plan is not history)', async () => {
    const programId = await programRepository.createProgram({ name: 'PPL', notes: null });
    const templateId = await programRepository.createTemplate(programId, benchTemplate);

    // A workout was performed from that template.
    const workoutId = await workoutRepository.saveCompletedWorkout({
      name: 'Push A',
      templateId,
      startedAt: null,
      endedAt: null,
      notes: null,
      exercises: [
        {
          exerciseId: BENCH,
          unilateralCounting: 'none',
          notes: null,
          sets: [{ weightKg: 100, reps: 5, rpe: null, warmup: false }],
        },
      ],
    });

    // Now the plan changes completely: different exercise, different targets.
    await programRepository.updateTemplate(templateId, {
      name: 'Push A (revised)',
      weekdays: [2, 4],
      notes: null,
      exercises: [
        {
          exerciseId: SQUAT,
          targetSets: 5,
          targetRepMin: 3,
          targetRepMax: 5,
          targetRpe: 9,
          restSeconds: 180,
          notes: null,
        },
      ],
    });

    // The performed workout is untouched — still bench, still 100×5.
    const workout = await workoutRepository.getById(workoutId);
    expect(workout?.name).toBe('Push A');
    expect(workout?.exercises).toHaveLength(1);
    expect(workout?.exercises[0]?.exerciseId).toBe(BENCH);
    expect(workout?.exercises[0]?.sets[0]).toMatchObject({ weightKg: 100, reps: 5 });
  });

  it('deleting a template keeps the workout and nulls its provenance (SET NULL)', async () => {
    const templateId = await programRepository.createTemplate(null, benchTemplate);
    const workoutId = await workoutRepository.saveCompletedWorkout({
      name: 'Push A',
      templateId,
      startedAt: null,
      endedAt: null,
      notes: null,
      exercises: [
        {
          exerciseId: BENCH,
          unilateralCounting: 'none',
          notes: null,
          sets: [{ weightKg: 100, reps: 5, rpe: null, warmup: false }],
        },
      ],
    });

    await programRepository.deleteTemplate(templateId);

    // Workout still exists (not cascade-deleted) …
    expect(await workoutRepository.getById(workoutId)).not.toBeNull();
    // … and its template_id was set to NULL, not left dangling.
    const row = testDb.sqlite
      .prepare('SELECT template_id FROM workouts WHERE id = ?')
      .get(workoutId) as { template_id: string | null };
    expect(row.template_id).toBeNull();
  });

  it('reports template-started workouts within the lookback for suggestions', async () => {
    const templateId = await programRepository.createTemplate(null, benchTemplate);
    await workoutRepository.saveCompletedWorkout({
      name: 'Push A',
      templateId,
      startedAt: null,
      endedAt: null,
      notes: null,
      exercises: [
        {
          exerciseId: BENCH,
          unilateralCounting: 'none',
          notes: null,
          sets: [{ weightKg: 100, reps: 5, rpe: null, warmup: false }],
        },
      ],
    });

    const uses = await programRepository.getRecentTemplateUses('2000-01-01');
    expect(uses).toHaveLength(1);
    expect(uses[0]?.templateId).toBe(templateId);
  });
});
