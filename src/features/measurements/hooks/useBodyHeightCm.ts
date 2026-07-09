import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { settingsRepository } from '@/data/repositories/settingsRepository';

/** The configured height in cm (Settings), for BMI derivation. Null when unset. */
export function useBodyHeightCm(): number | null {
  const version = useTableVersion('settings');
  const [heightCm, setHeightCm] = useState<number | null>(null);

  useEffect(() => {
    let live = true;
    void settingsRepository.get().then((s) => {
      if (live) setHeightCm(s.heightCm);
    });
    return () => {
      live = false;
    };
  }, [version]);

  return heightCm;
}
