import { createContext, useCallback, useState, type ReactNode } from 'react';
import { useColorScheme } from 'react-native';

import { prefs, type ThemeOverride } from '@/core/storage';

import { darkTheme, lightTheme, type Theme } from './tokens';

export const ThemeContext = createContext<Theme>(darkTheme);

interface ThemeControls {
  readonly override: ThemeOverride;
  readonly setOverride: (value: ThemeOverride) => void;
}

export const ThemeControlsContext = createContext<ThemeControls>({
  override: 'system',
  setOverride: () => undefined,
});

interface ThemeProviderProps {
  readonly children: ReactNode;
}

/**
 * Resolves the active theme: system-follow by default with a manual
 * dark/light/system override persisted in MMKV (DESIGN_SYSTEM §7).
 */
export function ThemeProvider(props: ThemeProviderProps): ReactNode {
  const scheme = useColorScheme();
  const [override, setOverrideState] = useState<ThemeOverride>(() => prefs.getThemeOverride());

  const setOverride = useCallback((value: ThemeOverride) => {
    prefs.setThemeOverride(value);
    setOverrideState(value);
  }, []);

  const resolved = override === 'system' ? scheme : override;
  const theme = resolved === 'light' ? lightTheme : darkTheme;

  return (
    <ThemeControlsContext.Provider value={{ override, setOverride }}>
      <ThemeContext.Provider value={theme}>{props.children}</ThemeContext.Provider>
    </ThemeControlsContext.Provider>
  );
}
