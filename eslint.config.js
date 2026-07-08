// ESLint configuration — rules defined by ARCHITECTURE §4 (boundaries) and
// DESIGN_SYSTEM §7 (token enforcement); config ownership per CODING_STANDARDS §3/§9.
const { defineConfig } = require('eslint/config');
const expoConfig = require('eslint-config-expo/flat');
const boundaries = require('eslint-plugin-boundaries');

module.exports = defineConfig([
  expoConfig,
  {
    ignores: ['dist/*', '.expo/*', 'node_modules/*', 'ios/*', 'android/*', 'scripts/*'],
  },
  {
    files: ['app/**/*.{ts,tsx}', 'src/**/*.{ts,tsx}'],
    plugins: { boundaries },
    settings: {
      'import/resolver': {
        typescript: { project: './tsconfig.json' },
      },
      'boundaries/elements': [
        { type: 'app', pattern: 'app', mode: 'folder' },
        { type: 'core-utils', pattern: 'src/core/utils', mode: 'folder' },
        { type: 'core', pattern: 'src/core', mode: 'folder' },
        { type: 'domain', pattern: 'src/domain', mode: 'folder' },
        { type: 'data', pattern: 'src/data', mode: 'folder' },
        { type: 'features', pattern: 'src/features/*', mode: 'folder', capture: ['featureName'] },
      ],
      'boundaries/dependency-nodes': ['import', 'export', 'dynamic-import'],
    },
    rules: {
      // ARCHITECTURE §4: one-way dependency rule; features never import features.
      'boundaries/element-types': [
        'error',
        {
          default: 'disallow',
          rules: [
            { from: 'app', allow: ['features', 'core', 'core-utils'] },
            {
              from: 'features',
              allow: [
                ['features', { featureName: '${from.featureName}' }],
                'domain',
                'data',
                'core',
                'core-utils',
              ],
            },
            { from: 'data', allow: ['data', 'domain', 'core', 'core-utils'] },
            { from: 'domain', allow: ['domain', 'core-utils'] },
            { from: 'core', allow: ['core', 'core-utils'] },
            { from: 'core-utils', allow: ['core-utils'] },
          ],
        },
      ],
      // ARCHITECTURE rule 2: domain (and pure utils) never import frameworks.
      'boundaries/external': [
        'error',
        {
          default: 'allow',
          rules: [
            {
              from: ['domain', 'core-utils'],
              disallow: ['react', 'react-*', 'expo', 'expo-*', '@expo/*', 'zustand'],
              message:
                'domain/ and core/utils are pure — no React/React Native/Expo imports (ARCHITECTURE §4).',
            },
          ],
        },
      ],
    },
  },
  {
    // DESIGN_SYSTEM §7: no raw hex colors or raw style numbers outside core/theme + core/ui.
    files: ['app/**/*.{ts,tsx}', 'src/features/**/*.{ts,tsx}', 'src/domain/**/*.{ts,tsx}'],
    rules: {
      'no-restricted-syntax': [
        'error',
        {
          selector: 'Literal[value=/^#(?:[0-9a-fA-F]{3,8})$/]',
          message:
            'Raw hex colors are forbidden outside core/theme and core/ui — use a theme token (DESIGN_SYSTEM §7).',
        },
        {
          selector:
            "Property[key.name=/^(fontSize|padding|paddingHorizontal|paddingVertical|margin|marginHorizontal|marginVertical|borderRadius)$/][value.type='Literal'][value.raw=/^[0-9]/]",
          message:
            'Raw style numbers are forbidden outside core/theme and core/ui — use spacing/type/radius tokens (DESIGN_SYSTEM §7).',
        },
      ],
    },
  },
]);
