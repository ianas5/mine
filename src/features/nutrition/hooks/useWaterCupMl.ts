import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { settingsRepository } from '@/data/repositories/settingsRepository';

const DEFAULT_CUP_ML = 250;

/** The configured water-cup size in ml (Settings), reactive. Defaults to 250 ml. */
export function useWaterCupMl(): number {
  const version = useTableVersion('settings');
  const [cupMl, setCupMl] = useState(DEFAULT_CUP_ML);

  useEffect(() => {
    let live = true;
    void settingsRepository.get().then((s) => {
      if (live) setCupMl(s.waterCupMl);
    });
    return () => {
      live = false;
    };
  }, [version]);

  return cupMl;
}
