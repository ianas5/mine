/**
 * The progress-photo file store (ARCHITECTURE §12): image bytes live on the
 * filesystem, keyed by a relative `fileName` under `<documentDirectory>/photos/`.
 * Abstracted behind this interface so the repository's file/row transaction is
 * testable with an in-memory double — the real expo-file-system impl is wired at
 * the composition root, exactly like the SQLite handle.
 */
export interface PhotoStore {
  /** Copies a picked/captured image into the photos dir under `fileName`. */
  saveFrom(sourceUri: string, fileName: string): Promise<void>;
  /** Removes the file if present (no-op when absent). */
  remove(fileName: string): void;
  exists(fileName: string): boolean;
  /** All image file names currently in the photos dir (for the orphan sweep). */
  listFileNames(): string[];
  /** Absolute URI for rendering (`<Image source={{ uri }}>`). */
  uri(fileName: string): string;
}

let current: PhotoStore | null = null;

export function getPhotoStore(): PhotoStore {
  if (current === null) {
    throw new Error('Photo store not initialized — the composition root must set it.');
  }
  return current;
}

/** Wire the store: the real expo store at startup, or an in-memory double in tests. */
export function setPhotoStore(store: PhotoStore): void {
  current = store;
}
