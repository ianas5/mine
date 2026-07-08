import { useContext } from 'react';

import { ThemeControlsContext } from './ThemeProvider';
import type { ThemeOverride } from '@/core/storage';

interface ThemeControls {
  readonly override: ThemeOverride;
  readonly setOverride: (value: ThemeOverride) => void;
}

/** Read/set the manual theme override (system | dark | light). */
export function useThemeControls(): ThemeControls {
  return useContext(ThemeControlsContext);
}
