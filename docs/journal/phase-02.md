# Phase 2 — Data Core

**Closed:** 2026-07-08 · **Tag:** `v0.3.0-phase2` · **Milestone:** M0 Foundation

## What was built

The SQLite/Drizzle persistence foundation and the first end-to-end data slice:

- **Schema & migration 0000:** Drizzle TypeScript schema (`data/schema/tables.ts`)
  with the `settings` table exactly per DATABASE §3.1 (single-row CHECK id=1,
  defaults 4 / 250ml); `drizzle-kit generate` migration committed
  (`0000_misty_maverick.sql` + expo runtime bundle); metro/babel wired for
  `.sql` inlining and the web driver's `.wasm` asset.
- **`core/db`:** connection with DATABASE §1 PRAGMAs (WAL, foreign_keys,
  busy_timeout) via `initDb`/`getDb`/`setDbForTesting`; **DbGate** (runs
  migrations before any feature renders; splash while pending; calm error state
  on failure — no white screen); **change-bus** (`emitTableChanges`,
  `subscribeToTables`, `useTableVersion` via `useSyncExternalStore`).
- **`core/storage`:** typed MMKV wrapper (`prefs`) holding disposable
  preferences only — currently the theme override.
- **Theme override:** ThemeProvider now resolves system|dark|light with the
  override persisted in MMKV (DESIGN_SYSTEM §7); `useThemeControls` exported.
- **Domain/data slice:** `Settings` domain model → `settingsMapper` →
  `settingsRepository` (get-creates-row, partial update, emits `settings` as
  the last step).
- **Test harness:** `createTestDb` — better-sqlite3 in-memory DB running the
  identical generated migrations; 5 repository tests against real SQLite
  (defaults, partial-patch persistence, single-row CHECK enforcement,
  change-bus emission on write, silence on read).
- **Settings screen** behind the Dashboard gear: theme SegmentedControl,
  weekly-target Stepper (1–14), height + default-bodyweight inputs (commit on
  blur), skeletons while loading, optimistic updates via `useSettings`.
- **CI:** `db:check` (all migrations apply to a fresh DB) added to the pipeline
  and to `npm run check`; `db:generate` script added.
- **Lint lanes extended:** `core-db`/`core-storage` element types; SQL imports
  (`expo-sqlite`/`drizzle-orm`/`better-sqlite3`) confined to `data/` + `core/db`;
  `react-native-mmkv` confined to `core/storage` — each proven by a deliberate
  violation that failed lint and was removed.

## What changed

New: `drizzle.config.ts`, `metro.config.js`, `babel.config.js`,
`src/data/{schema,mappers,repositories,testing}`, `src/core/db`,
`src/core/storage`, `src/domain/models`, `src/features/settings`,
`app/settings.tsx`, `scripts/check-migrations.js`. Modified: root layout
(DbGate + migrations injection at the composition root), ThemeProvider,
theme barrel, DashboardScreen (gear), eslint config, CI workflow, jest setup
(MMKV + safe-area mocks), package scripts. No frozen document was modified.

## Screens affected

Settings (new), Dashboard (gear button added). Screenshots:
`screenshots/phase-02/db-gate-error-{dark,light}.png` — see limitations.

## Manual tests performed (and results)

| Test | Method | Result |
|---|---|---|
| Migrations apply from zero | `db:check` + harness in every repo test | ✅ |
| Default row created on first read; updates persist & merge | repo tests on real SQLite | ✅ |
| Single-row invariant | raw INSERT id=2 rejected by CHECK (test) | ✅ |
| Change-bus emits on write, not on read | repo tests | ✅ |
| SQL/MMKV lane lint | 2 deliberate violations → both fail lint → removed → clean | ✅ |
| DbGate error state renders calmly (no crash) | Web run (where sqlite is unavailable) rendered the designed error state, both themes | ✅ (inadvertent but real verification) |
| `npm run check` | typecheck + lint + format + 55 tests + db:check | ✅ green |
| Settings screen visual walk / persistence across relaunch on device | — | ⚠️ requires device (TD-001, CP-A) |

## Known limitations

1. **No Settings-screen screenshots:** expo-sqlite's synchronous API does not
   run on web (tried SharedArrayBuffer flag and a COOP/COEP-isolated server —
   the sync bridge still times out). Web is not a supported runtime
   (DEVELOPMENT_WORKFLOW §1), and adding web-only DB code for screenshots would
   violate P2/P21. The archive holds the (correctly rendered) DbGate error
   state; the real Settings visuals + relaunch-persistence walk join TD-001's
   on-device pass at CP-A.
2. Input plausibility bounds are basic (positive-number parse; stepper 1–14);
   formal Zod validation with the FITNESS_DOMAIN §8.13 constants module arrives
   in Phase 4 as planned.
3. DbGate's error card uses plain RN text (core/db cannot depend on the design
   system without inverting layers); acceptable for a failure surface.

## Technical debt introduced

None new. TD-001's CP-A checklist now explicitly includes: Settings screen
walk, theme override persistence, weekly-target relaunch persistence, and
delete-app → reinstall → defaults return.

## Retrospective

**What went well?** The repository tests run against the *actual generated
migration* on real SQLite — the single-row CHECK constraint was verified at the
database level, not mocked. The change-bus stayed as small as ARCHITECTURE
demanded (an emitter and one hook).

**What was harder than expected?** The Drizzle-on-Expo toolchain: `.sql`
imports need metro `sourceExts` + babel `inline-import`, the web driver needs
`.wasm` as an asset, and typed routes regenerate only from the dev server (a
stale `router.d.ts` failed typecheck for the new `/settings` route). The
react-compiler lint also rejected the sync-from-props effect in the settings
form — restructured to a lazily-seeded child component, which is genuinely
better.

**What should change before the next phase?** After adding any route, bounce
the dev server to regenerate typed routes *before* running typecheck. Web
screenshots end at the DB boundary from here on — plan on the DbGate error
state being the only web-capturable full-app state until a web strategy is ever
wanted (it is not, per P19's offline-native focus).
