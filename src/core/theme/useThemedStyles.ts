import { useMemo } from 'react';

import { useTheme } from './useTheme';
import type { Theme } from './tokens';

/**
 * Resolves a style factory against the active theme, cached per theme object
 * (there are exactly two), so components don't rebuild styles every render.
 */
export function useThemedStyles<T>(factory: (theme: Theme) => T): T {
  const theme = useTheme();
  return useMemo(() => resolveCached(factory, theme), [factory, theme]);
}

const cache = new WeakMap<(theme: Theme) => unknown, WeakMap<Theme, unknown>>();

function resolveCached<T>(factory: (theme: Theme) => T, theme: Theme): T {
  let byTheme = cache.get(factory);
  if (!byTheme) {
    byTheme = new WeakMap();
    cache.set(factory, byTheme);
  }
  let styles = byTheme.get(theme);
  if (styles === undefined) {
    styles = factory(theme);
    byTheme.set(theme, styles);
  }
  return styles as T;
}
