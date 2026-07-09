import { Directory, File, Paths } from 'expo-file-system';
import * as DocumentPicker from 'expo-document-picker';
import * as Sharing from 'expo-sharing';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

import type { ArchiveContents, ArchiveStore } from './archiveStore';

const DATA_ENTRY = 'data.json';
const PHOTO_PREFIX = 'photos/';

const photosDir = (): Directory => new Directory(Paths.document, 'photos');
const stagingDir = (): Directory => new Directory(Paths.cache, 'import-staging');
const stagingPhotosDir = (): Directory => new Directory(stagingDir(), 'photos');

function ensure(dir: Directory): Directory {
  if (!dir.exists) dir.create({ intermediates: true });
  return dir;
}

/**
 * The real archive store (DATABASE §6): fflate for zip pack/unpack over the SDK-57
 * `File`/`Directory` byte API, plus expo-sharing / expo-document-picker for the OS
 * surfaces. All zip work is in-memory `Uint8Array`; extracted photos are staged
 * under `<cache>/import-staging/` and committed into `<document>/photos/` only after
 * the DB replace succeeds.
 */
export const expoArchiveStore: ArchiveStore = {
  async pack({ dataJson, photoNames, fileName }) {
    const entries: Record<string, Uint8Array> = { [DATA_ENTRY]: strToU8(dataJson) };
    const dir = photosDir();
    for (const name of photoNames) {
      const file = new File(dir, name);
      if (file.exists) entries[`${PHOTO_PREFIX}${name}`] = await file.bytes();
    }

    const zipped = zipSync(entries);
    const out = new File(Paths.cache, fileName);
    if (out.exists) out.delete();
    out.create();
    out.write(zipped);
    return out.uri;
  },

  async share(uri) {
    if (await Sharing.isAvailableAsync()) {
      await Sharing.shareAsync(uri, { mimeType: 'application/zip' });
    }
  },

  async pick() {
    const result = await DocumentPicker.getDocumentAsync({
      type: ['application/zip', 'application/octet-stream', '*/*'],
      copyToCacheDirectory: true,
    });
    return result.canceled ? null : (result.assets[0]?.uri ?? null);
  },

  async open(uri): Promise<ArchiveContents> {
    const bytes = await new File(uri).bytes();
    const unzipped = unzipSync(bytes);

    const dataBytes = unzipped[DATA_ENTRY];
    if (!dataBytes) throw new Error('archive is missing data.json');
    const dataJson = strFromU8(dataBytes);

    // Re-stage the extracted photos into a clean staging dir.
    const staged = ensure(stagingPhotosDir());
    for (const entry of staged.list()) entry.delete();

    const photoNames: string[] = [];
    for (const [path, content] of Object.entries(unzipped)) {
      if (!path.startsWith(PHOTO_PREFIX) || path === PHOTO_PREFIX) continue;
      const name = path.slice(PHOTO_PREFIX.length);
      const file = new File(staged, name);
      file.create();
      file.write(content);
      photoNames.push(name);
    }

    return { dataJson, photoNames };
  },

  async commitPhotos(photoNames) {
    const dest = ensure(photosDir());
    const staged = stagingPhotosDir();
    for (const name of photoNames) {
      const src = new File(staged, name);
      if (!src.exists) continue;
      const target = new File(dest, name);
      if (target.exists) target.delete();
      await src.copy(target);
    }
  },

  async cleanup() {
    const dir = stagingDir();
    if (dir.exists) dir.delete();
  },
};
