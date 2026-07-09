/**
 * Every way an import can be refused (DATABASE §6). Each code maps to a specific,
 * user-facing message; the common contract is that a thrown `ImportError` means
 * **nothing was written** — existing data is untouched. The UI never has to guess
 * whether a failure was safe.
 */
export type ImportErrorCode =
  | 'unreadable-archive' // could not open/extract the zip or find data.json
  | 'invalid-data' // JSON/Zod validation failed — malformed or corrupt document
  | 'unsupported-format' // archive format version this app does not understand
  | 'schema-too-new' // exported by a newer app version — "update the app first"
  | 'unsupported-schema' // older schema with no upgrader path to current
  | 'aborted-no-safety'; // user declined to continue after a failed safety export

export class ImportError extends Error {
  readonly code: ImportErrorCode;

  constructor(code: ImportErrorCode, message?: string) {
    super(message ?? code);
    this.name = 'ImportError';
    this.code = code;
  }
}
