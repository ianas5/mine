import { useCallback, useEffect, useState } from 'react';

import { useTableVersion } from '@/core/db';
import { settingsRepository } from '@/data/repositories/settingsRepository';
import type { Settings } from '@/domain/models';

interface UseSettingsResult {
  readonly settings: Settings | null;
  readonly update: (patch: Partial<Settings>) => void;
}

/** Settings read + optimistic update, re-queried on change-bus writes (ARCHITECTURE §7). */
export function useSettings(): UseSettingsResult {
  const version = useTableVersion('settings');
  const [settings, setSettings] = useState<Settings | null>(null);

  useEffect(() => {
    let live = true;
    void settingsRepository.get().then((value) => {
      if (live) setSettings(value);
    });
    return () => {
      live = false;
    };
  }, [version]);

  const update = useCallback((patch: Partial<Settings>) => {
    // Optimistic (P15): reflect immediately; the change-bus re-query confirms.
    setSettings((current) => (current ? { ...current, ...patch } : current));
    void settingsRepository.update(patch);
  }, []);

  return { settings, update };
}
