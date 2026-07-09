/** @jest-environment node */
import { setDbForTesting } from '@/core/db';
import type { PhotoAngle } from '@/domain/photos';

import { setPhotoStore, type PhotoStore } from '../photos/photoStore';
import { createTestDb, type TestDb } from '../testing/createTestDb';
import { photoRepository, type NewPhotoInput } from './photoRepository';

function makeFakeStore(): { store: PhotoStore; files: Map<string, string> } {
  const files = new Map<string, string>();
  const store: PhotoStore = {
    saveFrom: async (sourceUri, fileName) => {
      files.set(fileName, sourceUri);
    },
    remove: (fileName) => {
      files.delete(fileName);
    },
    exists: (fileName) => files.has(fileName),
    listFileNames: () => [...files.keys()],
    uri: (fileName) => `mem://${fileName}`,
  };
  return { store, files };
}

const input = (angle: PhotoAngle = 'front'): NewPhotoInput => ({
  date: '2026-07-09',
  angle,
  sourceUri: 'file:///tmp/pick.jpg',
  width: 1080,
  height: 1920,
  notes: null,
});

describe('photoRepository file/row lifecycle (real SQLite + in-memory store)', () => {
  let testDb: TestDb;
  let files: Map<string, string>;

  beforeEach(() => {
    testDb = createTestDb();
    setDbForTesting(testDb.db);
    const fake = makeFakeStore();
    setPhotoStore(fake.store);
    files = fake.files;
  });

  afterEach(() => testDb.close());

  it('writes the file and the row together (file-first, row-second)', async () => {
    await photoRepository.savePhoto(input());
    expect(files.size).toBe(1);
    const photos = await photoRepository.listPhotos();
    expect(photos).toHaveLength(1);
    expect(photos[0]?.fileMissing).toBe(false);
    expect(photos[0]?.uri).toMatch(/^mem:\/\//);
  });

  it('deletes the just-written file when the row insert fails (no orphan)', async () => {
    // An invalid angle violates the CHECK constraint → the insert throws.
    await expect(photoRepository.savePhoto(input('sideways' as PhotoAngle))).rejects.toBeDefined();

    expect(files.size).toBe(0); // file was rolled back
    expect(await photoRepository.listPhotos()).toHaveLength(0); // no row
  });

  it('removes the row then the file on delete', async () => {
    const id = await photoRepository.savePhoto(input());
    await photoRepository.deletePhoto(id);

    expect(files.size).toBe(0);
    expect(await photoRepository.listPhotos()).toHaveLength(0);
  });

  it('flags a row whose file has gone missing (renders a placeholder)', async () => {
    const id = await photoRepository.savePhoto(input());
    const fileName = `2026-07-09_front_${id}`;
    // Simulate the file vanishing out from under the row.
    files.delete([...files.keys()][0]!);

    const photos = await photoRepository.listPhotos();
    expect(photos[0]?.fileMissing).toBe(true);
    expect(fileName).toContain('front'); // name shape sanity
  });

  it('sweeps orphan files (no row) and counts rows with a missing file', async () => {
    await photoRepository.savePhoto(input());
    // An orphan file with no metadata row — e.g. a kill between file-write and insert.
    files.set('2026-07-09_back_orphan.jpg', 'file:///tmp/orphan.jpg');

    const result = await photoRepository.sweepOrphans();
    expect(result.removedFiles).toBe(1); // the orphan is gone
    expect(files.has('2026-07-09_back_orphan.jpg')).toBe(false);
    expect(result.missingRows).toBe(0); // the real photo's file is intact
    expect(files.size).toBe(1);
  });
});
