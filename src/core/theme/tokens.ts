/**
 * Design tokens (v1 stubs) — values per DESIGN_SYSTEM §2.2 (color), §3 (type),
 * §4 (space/radius/motion). Phase 0 ships the token *structure*; the full
 * primitive catalog that consumes it arrives in Phase 1.
 */

export interface ColorTokens {
  readonly bg: string;
  readonly surface: string;
  readonly surfaceRaised: string;
  readonly border: string;
  readonly textPrimary: string;
  readonly textSecondary: string;
  readonly textTertiary: string;
  readonly accent: string;
  readonly accentSoft: string;
  readonly positive: string;
  readonly attention: string;
  readonly danger: string;
  readonly chartLine: string;
  readonly chartMuted: string;
}

export interface TypeToken {
  readonly fontSize: number;
  readonly fontWeight: '400' | '500' | '600' | '700';
  readonly fontFamily: string;
}

export interface Theme {
  readonly mode: 'dark' | 'light';
  readonly color: ColorTokens;
  readonly space: typeof space;
  readonly radius: typeof radius;
  readonly type: typeof type;
  readonly motion: typeof motion;
}

export const space = {
  xs: 4,
  sm: 8,
  md: 12,
  lg: 16,
  xl: 20,
  xxl: 24,
  xxxl: 32,
} as const;

export const radius = {
  sm: 8,
  md: 12,
  lg: 16,
  full: 9999,
} as const;

export const motion = {
  fast: 150,
  base: 250,
  slow: 400,
} as const;

const inter = {
  regular: 'Inter_400Regular',
  medium: 'Inter_500Medium',
  semibold: 'Inter_600SemiBold',
  bold: 'Inter_700Bold',
} as const;

export const type = {
  display: { fontSize: 34, fontWeight: '700', fontFamily: inter.bold },
  title: { fontSize: 24, fontWeight: '700', fontFamily: inter.bold },
  heading: { fontSize: 18, fontWeight: '600', fontFamily: inter.semibold },
  body: { fontSize: 15, fontWeight: '400', fontFamily: inter.regular },
  bodyStrong: { fontSize: 15, fontWeight: '600', fontFamily: inter.semibold },
  caption: { fontSize: 13, fontWeight: '500', fontFamily: inter.medium },
  micro: { fontSize: 11, fontWeight: '600', fontFamily: inter.semibold },
} as const satisfies Record<string, TypeToken>;

export const darkColors: ColorTokens = {
  bg: '#0B0C0E',
  surface: '#151719',
  surfaceRaised: '#1D2023',
  border: '#26292E',
  textPrimary: '#F2F3F5',
  textSecondary: '#9BA0A8',
  textTertiary: '#5C6066',
  accent: '#FF6A3D',
  accentSoft: 'rgba(255, 106, 61, 0.12)',
  positive: '#3ECF8E',
  attention: '#FFB020',
  danger: '#FF5C5C',
  chartLine: '#FF6A3D',
  chartMuted: '#3A3E44',
};

export const lightColors: ColorTokens = {
  bg: '#F7F7F8',
  surface: '#FFFFFF',
  surfaceRaised: '#FFFFFF',
  border: '#E4E5E9',
  textPrimary: '#17181A',
  textSecondary: '#5D6167',
  textTertiary: '#9A9EA5',
  accent: '#E85320',
  accentSoft: 'rgba(232, 83, 32, 0.10)',
  positive: '#1F9D66',
  attention: '#B97A00',
  danger: '#D64545',
  chartLine: '#E85320',
  chartMuted: '#D8DADF',
};

export const darkTheme: Theme = {
  mode: 'dark',
  color: darkColors,
  space,
  radius,
  type,
  motion,
};

export const lightTheme: Theme = {
  mode: 'light',
  color: lightColors,
  space,
  radius,
  type,
  motion,
};
