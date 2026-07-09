import { useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { bodyRepository } from '@/data/repositories/bodyRepository';
import { settingsRepository } from '@/data/repositories/settingsRepository';

/**
 * The bodyweight (kg) used for bodyweight-load volume. FITNESS_DOMAIN §5.2 makes the
 * **latest weigh-in** canonical, so this prefers the most recent `body_snapshots`
 * weight and falls back to the manual `settings.defaultBodyweightKg` (then null).
 * Reactive to both body and settings writes (TD-009).
 */
export function useDefaultBodyweight(): number | null {
  const version = useTableVersion('settings', 'body');
  const [bodyweightKg, setBodyweightKg] = useState<number | null>(null);

  useEffect(() => {
    let live = true;
    void Promise.all([bodyRepository.getLatestWeightKg(), settingsRepository.get()]).then(
      ([latestWeighIn, settings]) => {
        if (live) setBodyweightKg(latestWeighIn ?? settings.defaultBodyweightKg);
      },
    );
    return () => {
      live = false;
    };
  }, [version]);

  return bodyweightKg;
}
