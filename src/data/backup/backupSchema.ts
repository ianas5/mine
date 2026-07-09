import { z } from 'zod';

/**
 * The backup document contract (DATABASE §6). Zod is the single validation gate on
 * import — the highest-risk edge (ARCHITECTURE §10) — so these schemas mirror the
 * table columns exactly (camelCase, matching Drizzle's select/insert shape). Unknown
 * keys are stripped (Zod default), so only known columns are ever inserted; a missing
 * or mistyped field fails the parse and aborts the import untouched.
 */

export const BACKUP_APP = 'personal-fitness-tracker';
/** Backup *archive* format version (independent of the DB `schemaVersion`). */
export const BACKUP_FORMAT = 1;

const int = z.number().int();
const real = z.number();
const str = z.string();
const flag = z.number().int().min(0).max(1); // integer-encoded boolean (0/1)

const settingsRow = z.object({
  // Always 1 (single-row table); the `settings_single_row` CHECK enforces it on insert.
  id: int,
  weeklyWorkoutTarget: int,
  defaultBodyweightKg: real.nullable(),
  heightCm: real.nullable(),
  waterCupMl: int,
  createdAt: int,
  updatedAt: int,
});

const exerciseRow = z.object({
  id: str,
  name: str,
  primaryMuscleGroup: str,
  secondaryMuscleGroups: str,
  loadType: str,
  defaultUnilateral: flag,
  isCustom: flag,
  isArchived: flag,
  notes: str.nullable(),
  createdAt: int,
  updatedAt: int,
});

const programRow = z.object({
  id: str,
  name: str,
  notes: str.nullable(),
  isActive: flag,
  isArchived: flag,
  createdAt: int,
  updatedAt: int,
});

const templateRow = z.object({
  id: str,
  programId: str.nullable(),
  name: str,
  position: int,
  weekday: int.nullable(),
  notes: str.nullable(),
  isArchived: flag,
  createdAt: int,
  updatedAt: int,
});

const templateExerciseRow = z.object({
  id: str,
  templateId: str,
  exerciseId: str,
  position: int,
  targetSets: int.nullable(),
  targetRepMin: int.nullable(),
  targetRepMax: int.nullable(),
  targetRpe: real.nullable(),
  restSeconds: int.nullable(),
  notes: str.nullable(),
});

const workoutRow = z.object({
  id: str,
  date: str,
  name: str,
  templateId: str.nullable(),
  startedAt: int.nullable(),
  endedAt: int.nullable(),
  notes: str.nullable(),
  createdAt: int,
  updatedAt: int,
});

const workoutExerciseRow = z.object({
  id: str,
  workoutId: str,
  exerciseId: str,
  position: int,
  unilateralCounting: str,
  notes: str.nullable(),
});

const setRow = z.object({
  id: str,
  workoutExerciseId: str,
  position: int,
  weightKg: real,
  reps: int,
  rpe: real.nullable(),
  rir: int.nullable(),
  isWarmup: flag,
  notes: str.nullable(),
});

const foodRow = z.object({
  id: str,
  name: str,
  servingAmount: real,
  servingUnit: str,
  kcal: int,
  proteinG: real,
  carbG: real,
  fatG: real,
  isQuickMeal: flag,
  isCustom: flag,
  isArchived: flag,
  createdAt: int,
  updatedAt: int,
});

const mealEntryRow = z.object({
  id: str,
  date: str,
  slot: str.nullable(),
  foodId: str.nullable(),
  foodName: str,
  loggedAmount: real,
  loggedUnit: str,
  kcal: int,
  proteinG: real,
  carbG: real,
  fatG: real,
  loggedAt: int,
});

const nutritionTargetRow = z.object({
  id: str,
  effectiveFrom: str,
  kcal: int,
  proteinG: real,
  carbG: real,
  fatG: real,
  waterMl: int.nullable(),
  createdAt: int,
  updatedAt: int,
});

const waterDayRow = z.object({
  date: str,
  ml: int,
  updatedAt: int,
});

const bodySnapshotRow = z.object({
  date: str,
  weightKg: real.nullable(),
  bodyFatPct: real.nullable(),
  muscleMassKg: real.nullable(),
  visceralFat: real.nullable(),
  bmi: real.nullable(),
  neckCm: real.nullable(),
  chestCm: real.nullable(),
  waistCm: real.nullable(),
  hipsCm: real.nullable(),
  leftArmCm: real.nullable(),
  rightArmCm: real.nullable(),
  leftForearmCm: real.nullable(),
  rightForearmCm: real.nullable(),
  leftThighCm: real.nullable(),
  rightThighCm: real.nullable(),
  leftCalfCm: real.nullable(),
  rightCalfCm: real.nullable(),
  createdAt: int,
  updatedAt: int,
});

const progressPhotoRow = z.object({
  id: str,
  date: str,
  angle: str,
  fileName: str,
  width: int.nullable(),
  height: int.nullable(),
  notes: str.nullable(),
  createdAt: int,
});

/** The `data` block: every backed-up table (all tables except `workout_drafts`, §6). */
export const backupDataSchema = z.object({
  settings: settingsRow,
  exercises: z.array(exerciseRow),
  programs: z.array(programRow),
  templates: z.array(templateRow),
  templateExercises: z.array(templateExerciseRow),
  workouts: z.array(workoutRow),
  workoutExercises: z.array(workoutExerciseRow),
  sets: z.array(setRow),
  foods: z.array(foodRow),
  mealEntries: z.array(mealEntryRow),
  nutritionTargets: z.array(nutritionTargetRow),
  waterDays: z.array(waterDayRow),
  bodySnapshots: z.array(bodySnapshotRow),
  progressPhotos: z.array(progressPhotoRow),
});
export type BackupData = z.infer<typeof backupDataSchema>;

/**
 * The archive header, parsed FIRST with lenient version fields so version
 * reconciliation can produce a precise error ("too new" / "unsupported") before
 * the full `data` validation runs against the current shape.
 */
export const backupHeaderSchema = z.object({
  app: str,
  format: int,
  schemaVersion: int,
  exportedAt: str,
  data: z.unknown(),
});
export type BackupHeader = z.infer<typeof backupHeaderSchema>;

export interface BackupEnvelope {
  readonly app: string;
  readonly format: number;
  readonly schemaVersion: number;
  readonly exportedAt: string;
  readonly data: BackupData;
}
