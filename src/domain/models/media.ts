import type { IsoDate } from '@/core/utils';
import type { PhotoAngle } from '@/domain/photos';

/** Progress-photo metadata (DATABASE §3.6). Image bytes live on the filesystem. */
export interface ProgressPhoto {
  readonly id: string;
  readonly date: IsoDate;
  readonly angle: PhotoAngle;
  /** Relative name under `<documentDirectory>/photos/`. */
  readonly fileName: string;
  readonly width: number | null;
  readonly height: number | null;
  readonly notes: string | null;
}
