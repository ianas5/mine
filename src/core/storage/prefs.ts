import { MMKV } from 'react-native-mmkv';

/**
 * Typed MMKV wrapper — DISPOSABLE preferences and cache only (ARCHITECTURE §3/§6).
 * Anything backup-worthy or history-affecting belongs in SQLite, never here.
 * Losing this store must never lose data — only a preference resets.
 */
const store = new MMKV({ id: 'prefs' });

export type ThemeOverride = 'system' | 'dark' | 'light';

const THEME_OVERRIDES: readonly ThemeOverride[] = ['system', 'dark', 'light'];

export const prefs = {
  getThemeOverride(): ThemeOverride {
    const raw = store.getString('themeOverride');
    return THEME_OVERRIDES.includes(raw as ThemeOverride) ? (raw as ThemeOverride) : 'system';
  },
  setThemeOverride(value: ThemeOverride): void {
    store.set('themeOverride', value);
  },
} as const;
