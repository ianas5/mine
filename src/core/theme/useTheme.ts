import { useContext } from 'react';

import { ThemeContext } from './ThemeProvider';
import type { Theme } from './tokens';

export function useTheme(): Theme {
  return useContext(ThemeContext);
}
