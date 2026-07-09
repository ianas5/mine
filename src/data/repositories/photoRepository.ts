import { desc, eq } from 'drizzle-orm';

import { emitTableChanges, getDb } from '@/core/db';
import type { IsoDate } from '@/core/utils';
import type { PhotoAngle } from '@/domain/photos';
import type { ProgressPhoto } from '@/domain/models';

import { newId } from '../id';
import { getPhotoStore } from '../photos/photoStore';
import { progressPhotos } from '../schema/tables';

export interface NewPhotoInput {
  readonly date: IsoDate;
  readonly angle: PhotoAngle;
  /** The picked/captured image URI to copy into the photos dir. */
  readonly sourceUri: string;
  readonly width: number | null;
  readonly height: number | null;
  readonly notes: string | null;
}

export interface PhotoWithStatus extends ProgressPhoto {
  readonly uri: string;
  /** True when the metadata row exists but its file is gone (render a placeholder). */
  readonly fileMissing: boolean;
}

export interface SweepResult {
  readonly removedFiles: number;
  readonly missingRows: number;
}

function rowToPhoto(row: typeof progressPhotos.$inferSelect): ProgressPhoto {
  return {
    id: row.id,
    date: row.date,
    angle: row.angle as PhotoAngle,
    fileName: row.fileName,
    width: row.width,
    height: row.height,
    notes: row.notes,
  };
}

export const photoRepository = {
  /**
   * Saves a photo as a single logical transaction (DATABASE §3.6): write the file
   * FIRST, then insert the row; if the insert fails, delete the just-written file so
   * no orphan remains. A DB row therefore never points at a missing file.
   */
  async savePhoto(input: NewPhotoInput): Promise<string> {
    const store = getPhotoStore();
    const id = newId('photo');
    const fileName = `${input.date}_${input.angle}_${id}.jpg`;

    await store.saveFrom(input.sourceUri, fileName);
    try {
      await getDb().insert(progressPhotos).values({
        id,
        date: input.date,
        angle: input.angle,
        fileName,
        width: input.width,
        height: input.height,
        notes: input.notes,
        createdAt: Date.now(),
      });
    } catch (cause) {
      store.remove(fileName); // row failed → no orphan file
      throw cause;
    }

    emitTableChanges('photos');
    return id;
  },

  /** All photos, newest first, each annotated with its render URI + missing-file flag. */
  async listPhotos(): Promise<PhotoWithStatus[]> {
    const store = getPhotoStore();
    const rows = await getDb()
      .select()
      .from(progressPhotos)
      .orderBy(desc(progressPhotos.date), desc(progressPhotos.createdAt));
    return rows.map((row) => {
      const photo = rowToPhoto(row);
      return {
        ...photo,
        uri: store.uri(photo.fileName),
        fileMissing: !store.exists(photo.fileName),
      };
    });
  },

  /** Deletes a photo: remove the row FIRST, then its file (DATABASE §3.6). */
  async deletePhoto(id: string): Promise<void> {
    const store = getPhotoStore();
    const rows = await getDb()
      .select({ fileName: progressPhotos.fileName })
      .from(progressPhotos)
      .where(eq(progressPhotos.id, id));

    await getDb().delete(progressPhotos).where(eq(progressPhotos.id, id));
    const fileName = rows[0]?.fileName;
    if (fileName !== undefined) store.remove(fileName);

    emitTableChanges('photos');
  },

  /**
   * Startup reconciliation (DATABASE §3.6): delete files with no metadata row
   * (orphans, e.g. a kill between file-write and row-insert) and count rows whose
   * file is missing (rendered as a placeholder, never deleted here).
   */
  async sweepOrphans(): Promise<SweepResult> {
    const store = getPhotoStore();
    const rows = await getDb().select({ fileName: progressPhotos.fileName }).from(progressPhotos);
    const known = new Set(rows.map((r) => r.fileName));

    let removedFiles = 0;
    for (const fileName of store.listFileNames()) {
      if (!known.has(fileName)) {
        store.remove(fileName);
        removedFiles += 1;
      }
    }

    let missingRows = 0;
    for (const row of rows) {
      if (!store.exists(row.fileName)) missingRows += 1;
    }

    return { removedFiles, missingRows };
  },
} as const;
