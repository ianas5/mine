import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { settingsRepository } from '@/data/repositories/settingsRepository';

/** The configured default bodyweight (kg) for bodyweight-load volume, or null. */
export function useDefaultBodyweight(): number | null {
  const version = useTableVersion('settings');
  const [bodyweightKg, setBodyweightKg] = useState<number | null>(null);

  useEffect(() => {
    let live = true;
    void settingsRepository.get().then((settings) => {
      if (live) setBodyweightKg(settings.defaultBodyweightKg);
    });
    return () => {
      live = false;
    };
  }, [version]);

  return bodyweightKg;
}
