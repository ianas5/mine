import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { photoRepository, type PhotoWithStatus } from '@/data/repositories/photoRepository';

/** All progress photos (newest first, with render URI + missing-file flag), reactive. */
export function usePhotos(): PhotoWithStatus[] | undefined {
  const version = useTableVersion('photos');
  const [photos, setPhotos] = useState<PhotoWithStatus[] | undefined>(undefined);

  useEffect(() => {
    let live = true;
    void photoRepository.listPhotos().then((rows) => {
      if (live) setPhotos(rows);
    });
    return () => {
      live = false;
    };
  }, [version]);

  return photos;
}
