import { Directory, File, Paths } from 'expo-file-system';

import type { PhotoStore } from './photoStore';

const photosDir = (): Directory => new Directory(Paths.document, 'photos');

function ensureDir(): Directory {
  const dir = photosDir();
  if (!dir.exists) dir.create({ intermediates: true });
  return dir;
}

/** The real filesystem-backed photo store (expo-file-system File/Directory API). */
export const expoPhotoStore: PhotoStore = {
  async saveFrom(sourceUri, fileName) {
    const dir = ensureDir();
    await new File(sourceUri).copy(new File(dir, fileName));
  },

  remove(fileName) {
    const file = new File(photosDir(), fileName);
    if (file.exists) file.delete();
  },

  exists(fileName) {
    return new File(photosDir(), fileName).exists;
  },

  listFileNames() {
    const dir = photosDir();
    if (!dir.exists) return [];
    return dir
      .list()
      .filter((entry): entry is File => entry instanceof File)
      .map((file) => file.name);
  },

  uri(fileName) {
    return new File(photosDir(), fileName).uri;
  },
};
