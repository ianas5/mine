import type { IsoDate } from '@/core/utils';

export const PHOTO_ANGLES = ['front', 'side', 'back'] as const;
export type PhotoAngle = (typeof PHOTO_ANGLES)[number];

export const PHOTO_ANGLE_LABELS: Record<PhotoAngle, string> = {
  front: 'Front',
  side: 'Side',
  back: 'Back',
};

interface AngleDated {
  readonly angle: PhotoAngle;
  readonly date: IsoDate;
}

/**
 * The smart-default capture angle (UI_UX §5.2): the **oldest missing** angle —
 * an angle never captured comes first (in front → side → back order), otherwise
 * the angle whose most-recent photo is the oldest. Pure over photo metadata.
 */
export function oldestMissingAngle(photos: readonly AngleDated[]): PhotoAngle {
  const latest = new Map<PhotoAngle, IsoDate>();
  for (const photo of photos) {
    const seen = latest.get(photo.angle);
    if (seen === undefined || photo.date > seen) latest.set(photo.angle, photo.date);
  }

  for (const angle of PHOTO_ANGLES) {
    if (!latest.has(angle)) return angle;
  }

  return PHOTO_ANGLES.reduce((oldest, angle) =>
    latest.get(angle)! < latest.get(oldest)! ? angle : oldest,
  );
}

export interface DatedPhotoGroup<T extends { readonly date: IsoDate }> {
  readonly date: IsoDate;
  readonly photos: readonly T[];
}

/** Groups photos by date, newest date first (the timeline gallery). */
export function groupPhotosByDate<T extends { readonly date: IsoDate }>(
  photos: readonly T[],
): DatedPhotoGroup<T>[] {
  const byDate = new Map<IsoDate, T[]>();
  for (const photo of photos) {
    const list = byDate.get(photo.date);
    if (list) list.push(photo);
    else byDate.set(photo.date, [photo]);
  }
  return [...byDate.entries()]
    .sort((a, b) => (a[0] < b[0] ? 1 : -1))
    .map(([date, list]) => ({ date, photos: list }));
}
