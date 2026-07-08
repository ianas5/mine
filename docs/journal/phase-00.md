# Phase 0 — Walking Skeleton

**Closed:** 2026-07-08 · **Tag:** `v0.1.0-phase0` · **Milestone:** M0 Foundation

## What was built

A running Expo (SDK 57) + React Native 0.86 + TypeScript (strict) project with the
complete toolchain and the five-tab shell, proving the pipeline end-to-end:

- Expo Router with root layout (Inter font loading via `@expo-google-fonts/inter`,
  splash-screen gate, theme provider) and a `(tabs)` group of five routes, each a
  thin delegator to a feature screen (ARCHITECTURE §5).
- `src/core/theme`: full token stubs (color dark+light per DESIGN_SYSTEM §2.2,
  space/radius/type/motion per §3–4), `ThemeProvider` (system-follow),
  `useTheme`, curated barrel. 4 token tests.
- Five placeholder screens (`src/features/*/screens/*Screen.tsx`) using tokens +
  Lucide icons only — no real features.
- ESLint flat config: `eslint-config-expo` + `typescript-eslint` +
  `eslint-plugin-boundaries` (ARCHITECTURE §4 layer rules incl. feature isolation
  and domain purity, resolved through the `@/` alias via
  `eslint-import-resolver-typescript`; import/export/dynamic-import all covered)
  + `no-restricted-syntax` token enforcement (DESIGN_SYSTEM §7). Zero-warning policy.
- Prettier (CODING_STANDARDS §9.1), Jest via `jest-expo`, `npm run check` gate,
  GitHub Actions CI (typecheck → lint → format check → test).
- TypeScript strict + `noUncheckedIndexedAccess`, `noImplicitOverride`,
  `noFallthroughCasesInSwitch`, `exactOptionalPropertyTypes`.

## What changed

New app scaffolding at repo root: `package.json`, `package-lock.json`, `app.json`,
`tsconfig.json`, `eslint.config.js`, `.prettierrc.json`, `.prettierignore`,
`.github/workflows/ci.yml`, `app/` (7 route files), `src/core/theme/` (5 files),
`src/features/*/screens/` (5 placeholder screens). No frozen document was modified.

## Screens affected

All five tab screens (Dashboard, Workouts, Nutrition, Measurements, Analytics) —
placeholder state, both themes. Screenshots: `screenshots/phase-00/` (10 files,
5 screens × dark/light).

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| App builds & bundles | `npx expo export --platform web` (all 5 routes exported) | ✅ |
| Cold start renders Dashboard | Chromium against static web build, iPhone-sized viewport (390×844@2x) | ✅ |
| Switch all 5 tabs | Real tab-bar clicks driven by Playwright | ✅ all five screens render with correct icon/title/copy and active-tab accent |
| Theme follows system dark/light | `colorScheme` emulation dark + light | ✅ tokens switch (verified visually in screenshots) |
| Boundary lint fails a deliberate violation | Cross-feature import, cross-feature `export…from`, and `react-native`-in-domain test files | ✅ all three error; removed after; lint clean |
| Token lint fails raw hex in a feature | Deliberate `#FF0000` literal | ✅ errors; removed after |
| `npm run check` | typecheck + lint (0 warnings) + format check + Jest | ✅ green, 4/4 tests pass |
| Kill & relaunch / physical iPhone launch | — | ⚠️ **Not performed** — see Known limitations |

## Known limitations

1. **Physical-device verification was not possible in this remote (Linux container)
   environment.** The acceptance criterion "builds and launches on a physical iPhone
   via dev build" is **outstanding**: verified instead via the static web build
   rendered at iPhone dimensions (react-native-web), which exercises the same
   routes, theme provider, fonts, and icons but is not a native runtime.
   `npx expo run:ios` must be run on a Mac + iPhone before CP-A (registered as debt).
2. Screenshots are web-build renders at iPhone size, not device captures. The
   web tab bar approximates but does not equal the native one (no safe-area inset).
3. The `Measurements` tab label truncates ("Measurem…") at the default tab-bar
   width — acceptable for placeholders; DESIGN_SYSTEM `TabBar` primitive (Phase 1)
   owns the final treatment.
4. Theme is system-follow only; the manual override arrives with MMKV in Phase 2
   (per plan, not a gap).
5. `db:generate` script and the CI migration check arrive in Phase 2 with the
   database layer (per plan).

## Technical debt introduced

Registered in `docs/TECH_DEBT.md`:

- **TD-001** — On-device (physical iPhone) launch verification outstanding.
- **TD-002** — Five placeholder screens duplicate their layout markup.

## Retrospective

**What went well?** The frozen documentation made every decision mechanical —
folder layout, token values, lint rules, and scripts were all transcription, not
design. The boundary lint catching all three violation classes (feature→feature,
export-from, framework-in-domain) on the first honest test after the resolver fix
validates the enforcement-first approach.

**What was harder than expected?** `eslint-plugin-boundaries` was silently inert
twice: first because the `@/` alias needed `eslint-import-resolver-typescript`,
then because `export … from` isn't covered by its default dependency nodes. A
"passing" lint proved nothing until a deliberate violation failed — the
acceptance criterion requiring a demonstrated failure caught what a green run
would have hidden. Also: `prettier --write .` reformatted the frozen markdown
documents; reverted and excluded `*.md` from Prettier's scope.

**What should change before the next phase?** (1) Always validate a new lint rule
with a deliberate violation before trusting it — adopt as standing practice for
every enforcement rule added in later phases. (2) Screenshot capture via the
static web build + Chromium works well; keep the capture script for future phases
but parameterize the phase number. (3) Physical-device verification needs the
owner's hardware — build/run instructions should be confirmed at CP-A.
