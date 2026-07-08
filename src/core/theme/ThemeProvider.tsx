import { createContext, type ReactNode } from 'react';
import { useColorScheme } from 'react-native';

import { darkTheme, lightTheme, type Theme } from './tokens';

export const ThemeContext = createContext<Theme>(darkTheme);

interface ThemeProviderProps {
  readonly children: ReactNode;
}

/**
 * Resolves the active theme. Phase 0: system-follow only; the manual
 * dark/light/system override (MMKV-backed, DESIGN_SYSTEM §7) arrives with
 * the storage layer in Phase 2.
 */
export function ThemeProvider(props: ThemeProviderProps): ReactNode {
  const scheme = useColorScheme();
  const theme = scheme === 'light' ? lightTheme : darkTheme;
  return <ThemeContext.Provider value={theme}>{props.children}</ThemeContext.Provider>;
}
