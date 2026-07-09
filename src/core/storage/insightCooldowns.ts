import { MMKV } from 'react-native-mmkv';

/**
 * Insight cooldown state (ANALYTICS §6.4/§8) — MMKV only, disposable. Maps an insight's
 * `instanceKey` to when it last fired + its classification (for flip detection). Losing
 * this merely re-shows a card; it is never a source of truth. Import/restore clears it.
 */
export interface CooldownEntry {
  readonly lastFired: string;
  readonly classification: string;
}
export type CooldownMap = Readonly<Record<string, CooldownEntry>>;

const store = new MMKV({ id: 'insight-cooldowns' });
const KEY = 'cooldowns';

export const insightCooldowns = {
  get(): CooldownMap {
    const raw = store.getString(KEY);
    if (!raw) return {};
    try {
      return JSON.parse(raw) as CooldownMap;
    } catch {
      return {};
    }
  },
  set(map: CooldownMap): void {
    store.set(KEY, JSON.stringify(map));
  },
  clear(): void {
    store.delete(KEY);
  },
} as const;
