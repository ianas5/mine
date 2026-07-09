import type { PhotoStore } from '../photos/photoStore';
import type { ArchiveContents, ArchiveStore } from './archiveStore';
import { BACKUP_APP, BACKUP_FORMAT, type BackupData } from './backupSchema';
import { CURRENT_SCHEMA_VERSION } from './schemaVersion';

/**
 * Shared test doubles + fixtures for the backup suites. In-memory fakes for the
 * photo store (a `Map` standing in for `<document>/photos/`) and the archive store
 * (a `Map` of packed archives sharing that photo map), plus one valid row per table
 * with consistent foreign keys. Not a `.test` file — imported by the backup tests.
 */

export function sampleBackupData(): BackupData {
  return {
    settings: {
      id: 1,
      weeklyWorkoutTarget: 4,
      defaultBodyweightKg: 80,
      heightCm: 180,
      targetWeightKg: 75,
      waterCupMl: 250,
      createdAt: 1,
      updatedAt: 1,
    },
    exercises: [
      {
        id: 'ex1',
        name: 'Bench Press',
        primaryMuscleGroup: 'chest',
        secondaryMuscleGroups: '["triceps"]',
        loadType: 'external',
        defaultUnilateral: 0,
        isCustom: 0,
        isArchived: 0,
        notes: null,
        createdAt: 1,
        updatedAt: 1,
      },
    ],
    programs: [
      {
        id: 'pr1',
        name: 'PPL',
        notes: null,
        isActive: 1,
        isArchived: 0,
        createdAt: 1,
        updatedAt: 1,
      },
    ],
    templates: [
      {
        id: 'tp1',
        programId: 'pr1',
        name: 'Push',
        position: 0,
        weekday: 0,
        notes: null,
        isArchived: 0,
        createdAt: 1,
        updatedAt: 1,
      },
    ],
    templateExercises: [
      {
        id: 'te1',
        templateId: 'tp1',
        exerciseId: 'ex1',
        position: 0,
        targetSets: 3,
        targetRepMin: 5,
        targetRepMax: 8,
        targetRpe: 8,
        restSeconds: 120,
        notes: null,
      },
    ],
    workouts: [
      {
        id: 'w1',
        date: '2026-01-01',
        name: 'Push',
        templateId: 'tp1',
        startedAt: 1000,
        endedAt: 2000,
        notes: null,
        createdAt: 1,
        updatedAt: 1,
      },
    ],
    workoutExercises: [
      {
        id: 'we1',
        workoutId: 'w1',
        exerciseId: 'ex1',
        position: 0,
        unilateralCounting: 'none',
        notes: null,
      },
    ],
    sets: [
      {
        id: 's1',
        workoutExerciseId: 'we1',
        position: 0,
        weightKg: 100,
        reps: 5,
        rpe: 8,
        rir: 2,
        isWarmup: 0,
        notes: null,
      },
    ],
    foods: [
      {
        id: 'f1',
        name: 'Chicken Breast',
        servingAmount: 100,
        servingUnit: 'g',
        kcal: 165,
        proteinG: 31,
        carbG: 0,
        fatG: 3.6,
        isQuickMeal: 0,
        isCustom: 1,
        isArchived: 0,
        createdAt: 1,
        updatedAt: 1,
      },
    ],
    mealEntries: [
      {
        id: 'm1',
        date: '2026-01-01',
        slot: 'lunch',
        foodId: 'f1',
        foodName: 'Chicken Breast',
        loggedAmount: 200,
        loggedUnit: 'g',
        kcal: 330,
        proteinG: 62,
        carbG: 0,
        fatG: 7.2,
        loggedAt: 1,
      },
    ],
    nutritionTargets: [
      {
        id: 'nt1',
        effectiveFrom: '2026-01-01',
        kcal: 2500,
        proteinG: 180,
        carbG: 250,
        fatG: 80,
        waterMl: 3000,
        createdAt: 1,
        updatedAt: 1,
      },
    ],
    waterDays: [{ date: '2026-01-01', ml: 2000, updatedAt: 1 }],
    bodySnapshots: [
      {
        date: '2026-01-01',
        weightKg: 80,
        bodyFatPct: 15,
        muscleMassKg: null,
        visceralFat: null,
        bmi: 24.7,
        neckCm: null,
        chestCm: null,
        waistCm: 82,
        hipsCm: null,
        leftArmCm: null,
        rightArmCm: null,
        leftForearmCm: null,
        rightForearmCm: null,
        leftThighCm: null,
        rightThighCm: null,
        leftCalfCm: null,
        rightCalfCm: null,
        createdAt: 1,
        updatedAt: 1,
      },
    ],
    progressPhotos: [
      {
        id: 'p1',
        date: '2026-01-01',
        angle: 'front',
        fileName: '2026-01-01_front_p1.jpg',
        width: 1080,
        height: 1440,
        notes: null,
        createdAt: 1,
      },
    ],
  };
}

/** JSON for an archive's `data.json`, with optional envelope-field overrides. */
export function makeEnvelopeJson(
  data: BackupData,
  overrides: Partial<{
    app: string;
    format: number;
    schemaVersion: number;
    exportedAt: string;
  }> = {},
): string {
  return JSON.stringify({
    app: BACKUP_APP,
    format: BACKUP_FORMAT,
    schemaVersion: CURRENT_SCHEMA_VERSION,
    exportedAt: '2026-01-02T00:00:00.000Z',
    data,
    ...overrides,
  });
}

/** In-memory `PhotoStore` over a shared `fileName → content` map (the "photos dir"). */
export function createFakePhotoStore(files: Map<string, string>): PhotoStore {
  return {
    async saveFrom(sourceUri, fileName) {
      files.set(fileName, `bytes:${sourceUri}`);
    },
    remove(fileName) {
      files.delete(fileName);
    },
    exists(fileName) {
      return files.has(fileName);
    },
    listFileNames() {
      return [...files.keys()];
    },
    uri(fileName) {
      return `file://photos/${fileName}`;
    },
  };
}

interface StoredArchive {
  readonly dataJson: string;
  readonly photos: Map<string, string>;
}

export interface FakeArchiveControls {
  failSafetyExport: boolean;
  pickUri: string | null;
  lastExportUri: string | null;
  lastSharedUri: string | null;
  readonly archives: Map<string, StoredArchive>;
  /** Directly inject an archive (for malformed/version tests). */
  put(uri: string, dataJson: string, photos?: Map<string, string>): void;
}

/** In-memory `ArchiveStore` sharing the photo-dir map; pack reads it, commit writes it. */
export function createFakeArchiveStore(files: Map<string, string>): {
  store: ArchiveStore;
  controls: FakeArchiveControls;
} {
  const archives = new Map<string, StoredArchive>();
  let staged: Map<string, string> = new Map();

  const controls: FakeArchiveControls = {
    failSafetyExport: false,
    pickUri: null,
    lastExportUri: null,
    lastSharedUri: null,
    archives,
    put(uri, dataJson, photos = new Map()) {
      archives.set(uri, { dataJson, photos });
    },
  };

  const store: ArchiveStore = {
    async pack({ dataJson, photoNames, fileName }) {
      if (fileName.startsWith('pre-import-safety') && controls.failSafetyExport) {
        throw new Error('safety export failed (simulated)');
      }
      const photos = new Map<string, string>();
      for (const name of photoNames) {
        const content = files.get(name);
        if (content !== undefined) photos.set(name, content);
      }
      const uri = `archive://${fileName}`;
      archives.set(uri, { dataJson, photos });
      if (fileName.startsWith('fitness-backup')) controls.lastExportUri = uri;
      return uri;
    },
    async share(uri) {
      controls.lastSharedUri = uri;
    },
    async pick() {
      return controls.pickUri;
    },
    async open(uri): Promise<ArchiveContents> {
      const archive = archives.get(uri);
      if (!archive) throw new Error(`no such archive: ${uri}`);
      staged = new Map(archive.photos);
      return { dataJson: archive.dataJson, photoNames: [...archive.photos.keys()] };
    },
    async commitPhotos(photoNames) {
      for (const name of photoNames) {
        const content = staged.get(name);
        if (content !== undefined) files.set(name, content);
      }
    },
    async cleanup() {
      staged = new Map();
    },
  };

  return { store, controls };
}
