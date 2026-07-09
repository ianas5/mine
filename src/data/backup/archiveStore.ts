/**
 * The backup archive I/O boundary (DATABASE §6): zip packing/unpacking, the OS
 * share sheet, the document picker, and staging of extracted photo files. Abstracted
 * behind this interface — exactly like `PhotoStore` — so the export/import
 * orchestration (validate → safety-export → replace → reconcile) is fully testable
 * with an in-memory double; the real expo/fflate implementation is wired at the
 * composition root.
 */
export interface ArchiveContents {
  /** The raw text of `data.json` inside the archive. */
  readonly dataJson: string;
  /** The photo file names staged from `photos/`, ready for `commitPhotos`. */
  readonly photoNames: readonly string[];
}

export interface ArchiveStore {
  /**
   * Packs `data.json` plus the named photos (read from the live photos dir) into a
   * `.zip` and returns its URI. Used for both the user export and the safety export.
   */
  pack(input: {
    readonly dataJson: string;
    readonly photoNames: readonly string[];
    readonly fileName: string;
  }): Promise<string>;
  /** Presents an archive to the user (OS share sheet). */
  share(uri: string): Promise<void>;
  /** Opens the document picker; resolves to the chosen archive URI, or null if cancelled. */
  pick(): Promise<string | null>;
  /** Extracts an archive to a staging area and returns its `data.json` + staged photo names. */
  open(uri: string): Promise<ArchiveContents>;
  /** Copies the named staged photos from the last `open` into the live photos dir. */
  commitPhotos(photoNames: readonly string[]): Promise<void>;
  /** Discards any staging/temp state (always safe to call). */
  cleanup(): Promise<void>;
}

let current: ArchiveStore | null = null;

export function getArchiveStore(): ArchiveStore {
  if (current === null) {
    throw new Error('Archive store not initialized — the composition root must set it.');
  }
  return current;
}

/** Wire the store: the real expo/fflate impl at startup, or an in-memory double in tests. */
export function setArchiveStore(store: ArchiveStore): void {
  current = store;
}
