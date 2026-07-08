# ARCHITECTURE.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Owner concern:** structural codebase decisions and technology responsibilities · **Depends on:** FITNESS_DOMAIN.md.
>
> This document owns **how the code is organized and which technology is responsible for what.** It does not define fitness meaning (FITNESS_DOMAIN), the SQL schema (DATABASE), metric interpretation (ANALYTICS_ENGINE), visual tokens (DESIGN_SYSTEM), screen flows (UI_UX_GUIDELINES), file-level style (CODING_STANDARDS), or process (DEVELOPMENT_WORKFLOW). Where those concerns are touched here, this document defines the **boundary and the seam**, and points to the owning document for the detail.

---

## 1. Architectural Goals

The architecture serves one long-lived personal product. It optimizes for, in priority order:

1. **Correctness of history** — the user's data is irreplaceable and must survive years of app changes.
2. **Local-first, always-offline** — the app is fully functional with no network, ever. There is no backend.
3. **Testable business logic** — the fitness/analytics logic (the app's actual value) is pure and unit-tested in isolation from React Native.
4. **Maintainability for a team of one** — clear boundaries, small files, no cross-cutting spaghetti, easy to return to after months away.
5. **Speed** — instant interactions on an iPhone; heavy computation kept off the render path.

A goal-shaped warning: this is a **personal app, not an enterprise system**. Every boundary below exists to serve the five goals above. When a rule would force ceremony without serving a goal, the rule bends (see §9.1 on domain pragmatism).

---

## 2. Architectural Principles

1. **Local-first.** SQLite on-device is the single source of truth for all durable data. The UI reads and writes locally and is never blocked on a network.
2. **Strict layering with a one-way dependency rule** (§4). Inner layers never know about outer layers.
3. **Pure domain core, pragmatically scoped.** All *shared* fitness math and analytics live in a framework-free `domain/` layer that imports **no** React, React Native, Expo, or SQLite. Feature-*specific* logic may live inside its feature (still pure, still separated from UI) — see §9.1.
4. **Feature-based organization.** Code is grouped by product feature (workouts, nutrition, …), not by technical type. Features are **isolated** — no feature imports another feature.
5. **Separation of UI and business logic.** Components render and capture input; they never compute domain metrics or touch persistence directly.
6. **Unidirectional data flow.** Input → validate → mutate persistence → notify → re-query → recompute → render. No back-channels.
7. **Strong typing end-to-end.** TypeScript `strict`; Zod validates every external boundary and its inferred types feed the domain.
8. **Derived data is never a source of truth.** PRs, trends, and insights are always recomputed from source records; any caching is disposable.

---

## 3. Technology Responsibility Matrix

Each technology has one job and explicit non-goals. Using a technology outside its lane is an architecture violation.

| Technology | Responsible for | Explicitly NOT for |
|---|---|---|
| **Expo (SDK, latest stable)** | App runtime, native module access (SQLite, FileSystem, Haptics, ImagePicker), build/OTA via EAS. | Business logic. |
| **Expo Router** | Navigation topology only — file-based routes, tab/stack/modal structure, params, deep-linking. | Holding app state or business logic; route files stay thin. |
| **React Native** | Rendering and user interaction (the presentation layer). | Computation, persistence, validation. |
| **TypeScript (strict)** | Static typing across every layer; domain-model types are canonical. | — (always on). |
| **SQLite (`expo-sqlite`)** | **Single source of truth** for all durable domain data + anything that must appear in a backup. Relational history, migrations. | Ephemeral UI state; large binaries (photos are files). |
| **Drizzle ORM + drizzle-kit** | Typed schema, typed queries, and generated SQL migrations, entirely inside `data/` (detail owned by DATABASE). | Any use outside `data/`; runtime schema mutation. |
| **MMKV** | Fast key-value for **disposable** UI preferences and cache we can lose without losing history (theme, active time-range, last tab, memoized-analytics cache). | Anything that belongs in a backup or affects historical correctness. |
| **Zustand** | **Global ephemeral session state** — the in-progress workout, rest-timer, cross-screen transient selections, and the reactive data-cache layer over repositories. | Durable data (that's SQLite); form-local state (that's RHF). |
| **React Hook Form** | **Form-local** state, field wiring, submission lifecycle for every input form. | Global state; persistence. |
| **Zod** | The **validation boundary** and schema source-of-truth at every external edge (user forms, backup import). Inferred types flow inward. | Runtime business rules beyond shape/range validation. |
| **Victory Native XL** | Chart rendering inside `ChartFrame` (detail owned by DESIGN_SYSTEM). | Computing, resampling, or interpreting analytics — it draws engine-bucketed data only. |
| **`expo-file-system`** | Storage of **progress-photo image files** in the app document directory; export/import file IO. | Storing image bytes in SQLite. |

*Note: Inter (typeface) and Lucide (icon set) are approved design assets, owned by DESIGN_SYSTEM.*

**The MMKV ↔ SQLite rule (memorize this):** *If losing it would lose history or belong in a backup → SQLite. If it's a preference or cache we can rebuild → MMKV.*

**Approved decision:** long-lived domain configuration — `weeklyWorkoutTarget`, `defaultBodyweight`, `height` — and **time-versioned nutrition targets** (`effectiveFrom` history, per FITNESS_DOMAIN §4) are **domain data → SQLite**. They appear in backups and can affect computed history. MMKV never holds them.

**Approved decision:** no data-fetching library (React Query etc.). Reactivity is the simple local change-bus in §7. Revisit only if a clear future need emerges (e.g. an optional sync backend), and only via an amendment to this document.

---

## 4. Layered Architecture & the Dependency Rule

Five layers, with a strict **one-way dependency direction**. An arrow means "may import from."

```
  app/  (routes)                        ← thinnest; navigation only
     │
     ▼
  features/  (screens, components, hooks, feature stores, feature logic)
     │            │
     ▼            ▼
  data/        domain/                   ← domain is a PURE leaf
  (repositories, SQL) │                    (no framework imports)
     │            │
     ▼            ▼
  core/  (db, storage, ui, theme, utils, config)
```

**The rules (enforced by lint — rule config owned by CODING_STANDARDS, defined here):**

- **`app/`** imports feature screens and navigation helpers only — plus **composition-root wiring in `app/_layout`**: mounting global providers and injecting data-layer artifacts that must not be reached by an inverted dependency (e.g. the `data/schema` migration bundle passed into the `core/db` DB-ready gate). This composition-root exception is confined to the root layout; ordinary route files stay thin (rule 5). Nothing imports from `app/`.
- **`features/*`** may import from `domain/`, `data/`, and `core/`. **A feature must never import from another feature.** Shared cross-feature code is promoted to `core/` (infra/ui) or `domain/` (logic).
- **`domain/`** may import only from `domain/` and pure `core/utils`. It must **not** import React, React Native, Expo, `data/`, `features/`, or any UI. It is a pure, framework-agnostic leaf.
- **`data/`** may import from `core/db` and `domain/models`. It contains the **only** SQL in the codebase. No UI, no features.
- **`core/`** is the foundation; it imports nothing above it. `core/ui` and `core/theme` are **owned by DESIGN_SYSTEM**; `core/db` schema detail is **owned by DATABASE**.

This is what guarantees principles #3 and #5: the valuable logic is isolated and testable, and the UI physically cannot reach past its hooks into SQL.

---

## 5. Project / Folder Structure (feature-based)

```
app/                         # Expo Router — routes are thin delegators
  _layout.tsx                # root providers (theme, db-ready gate)
  (tabs)/
    _layout.tsx              # bottom tab navigator (5 tabs)
    index.tsx                # Dashboard
    workouts/…               # stack per tab
    nutrition/…
    measurements/…
    analytics/…
  (modals)/                  # logging modals (add-weight, log-meal, …)

src/
  core/                      # cross-cutting foundation (imports nothing above)
    db/                      # sqlite connection, migration runner, change-bus
    storage/                 # mmkv wrappers (typed key namespaces)
    ui/                      # design-system PRIMITIVES        [owned by DESIGN_SYSTEM]
    theme/                   # tokens: color, type, spacing     [owned by DESIGN_SYSTEM]
    navigation/              # nav types & helpers
    validation/              # shared zod helpers/refinements
    utils/                   # date, math, number/format (pure)
    config/                  # constants, env, feature flags
    types/                   # shared base/util types

  domain/                    # PURE shared business logic (framework-free, unit-tested)
    models/                  # canonical entity TYPES (Workout, Set, BodySnapshot…)
    fitness/                 # formulas: volume, e1RM, PRs, consistency  [FITNESS_DOMAIN]
    analytics/               # metrics, trends, insight rules             [ANALYTICS_ENGINE]

  data/                      # persistence realization
    schema/                  # DDL + migrations                          [owned by DATABASE]
    repositories/            # one per aggregate; ONLY place SQL lives
    mappers/                 # DB row  <->  domain model
    seed/                    # seed exercise library, starter foods
    backup/                  # export/import serialization               [format by DATABASE]

  features/                  # UI + orchestration, grouped by product area
    dashboard/
    workouts/
      components/            # feature-specific components
      hooks/                 # data + interaction hooks (useWorkoutSession…)
      stores/                # zustand stores (active session, timer)
      screens/               # screen components; routes in app/ delegate here
      schemas/               # zod form schemas (log-set, edit-workout)
      logic/                 # feature-specific pure logic (optional; see §9.1)
    nutrition/
    measurements/
    photos/
    analytics/
```

Routes in `app/` contain no logic: each route file renders a screen component from `src/features/<x>/screens`. This keeps navigation swappable and screens testable in isolation.

---

## 6. State Ownership (the definitive table)

Every piece of state has exactly one home. When adding state, place it by this table.

| State category | Home | Persistence | Examples |
|---|---|---|---|
| Durable domain data | **SQLite** | in backups | workouts, sets, nutrition days, meal entries, foods, body snapshots, photo metadata, programs, templates, phases, **time-versioned nutrition targets** |
| Domain configuration | **SQLite** (settings) | in backups | `weeklyWorkoutTarget`, `defaultBodyweight`, `height` |
| Crash-safe session draft | **SQLite** (draft row) | not exported* | the active workout, checkpointed (§7.1) |
| Disposable prefs / cache | **MMKV** | not backed up | theme mode, active time-range, last active tab, cached analytics snapshots |
| Global ephemeral session | **Zustand** | in-memory | active workout session state, rest-timer, transient cross-screen selection |
| Reactive data cache | **Zustand** (per feature) | in-memory | loaded query results + invalidation (§7) |
| Form-local | **React Hook Form** | none | log-set form, add-meal form, measurement form |
| Component-local UI | `useState`/`useReducer` | none | expanded row, toggle, local scroll state |

\* Draft rows are recovery data, not history; exports exclude them (decided in DATABASE §6).

There is **no server-cache layer** — there is no server.

---

## 7. Data Flow Patterns

Because there is no backend and no data-fetching library, reactivity over SQLite uses a small **change-bus** in `core/db`: every write announces which table(s) changed; feature data-hooks subscribe and re-query. This keeps the UI in sync without polling. It is deliberately minimal — an emitter keyed by table name, nothing more (no query keys, no staleness policies, no retries; none of that is needed offline-only).

**Read path**
```
Component
  → feature hook (useX)              e.g. useWorkoutHistory()
    → feature store / query hook     (Zustand cache; subscribes to change-bus)
      → repository.query()           (data/) — the only SQL
        → SQLite → rows → mapper → domain models
    → (optional) domain/analytics compute (pure)
  → render (format/round here)
```

**Write path**
```
Form (RHF) → Zod.parse(input)        (reject → field errors, no write)
  → feature action (store/service)
    → repository.insert/update/delete (data/, wrapped in a transaction)
      → SQLite write
      → change-bus.emit(table)
        → subscribed hooks re-query → recompute derived → UI updates
```

### 7.1 Active workout session (special case) — approved

The in-progress workout lives in a **Zustand store** (fast, no DB churn per keystroke). Two guarantees:

1. **Crash safety (required, not optional):** the session is **checkpointed to a SQLite draft row** — on meaningful mutations (set completed, exercise added) debounced, and on app-background. On next launch, if a draft exists, the app offers **Resume / Discard**. A crash or force-close never loses a workout in progress.
2. **Single durable write on finish:** completing the workout writes it once, transactionally, via the workouts repository (and deletes the draft); PR/analytics recomputation is then triggered by the change-bus.

This isolates high-frequency editing from durable storage while making the session loss-proof.

---

## 8. Navigation Architecture (Expo Router)

- **Root layout** (`app/_layout.tsx`): mounts global providers (theme, safe-area) and a **DB-ready gate** that runs migrations/seed before rendering the app.
- **Tab group** (`app/(tabs)/_layout.tsx`): the five fixed bottom tabs from the vision — **Dashboard, Workouts, Nutrition, Measurements, Analytics**. Tab order and identity are fixed.
- **Per-tab stacks:** each tab is a stack for drill-down (e.g. Workouts → workout detail → exercise history).
- **Modals** (`app/(modals)/…`): the logging surfaces — Start/Log Workout, Log Meal, Add Weight, Add Measurements, Add Photo — presented modally so quick actions from any screen feel instant.
- **Params are typed;** IDs (not objects) are passed through routes, and screens re-query by ID. Detailed screen flows and hierarchy are **owned by UI_UX_GUIDELINES**; this document fixes only the topology (5 tabs + modal logging).

---

## 9. Domain & Analytics Placement

- `domain/fitness/` implements the **FITNESS_DOMAIN formulas** (volume, e1RM/Epley, PR detection, consistency, adherence) as **pure functions** over domain models. No I/O.
- `domain/analytics/` implements the **ANALYTICS_ENGINE** computations (time-series, regression/trend with deadbands, recomposition signal, insight-rule evaluation) — also pure, consuming data handed to it by repositories/hooks.
- `domain/models/` holds the **canonical TypeScript types** for entities. DB row types (in `data/`) and Zod input types map to/from these via `data/mappers`.
- Consumers (Dashboard, Analytics, Measurements) call these via feature hooks; they never inline formulas. Rounding/formatting for display happens in the UI/`core/utils`, never in `domain/` (which stays raw).

### 9.1 Pragmatism rule (anti-over-engineering) — approved

`domain/` is for **shared domain concepts**: anything defined in FITNESS_DOMAIN, anything two or more features consume, and anything the Dashboard/Analytics aggregate across pillars. It is **not** a mandatory home for every scrap of logic:

- **Feature-specific pure logic may live inside its feature** (e.g. `features/workouts/logic/` — set-row reordering rules, rest-timer sequencing, template-to-session expansion). Same purity discipline (no I/O, no framework imports, testable), just located where it's used.
- **Promotion, not prediction:** logic starts in its feature and is promoted to `domain/` the moment a second consumer appears or it turns out to encode a FITNESS_DOMAIN rule. Do not build speculative abstractions in `domain/` for single-consumer logic.
- **The line that never bends:** FITNESS_DOMAIN formulas (volume, e1RM, PRs, adherence, trends, recomposition) always live in `domain/` — they are the app's canon and must have exactly one implementation.

This keeps the domain layer honest and small instead of becoming a ceremonial dumping ground.

---

## 10. Validation Strategy (Zod)

- **Zod schemas are the source of truth at every external boundary:** user forms (with React Hook Form via a resolver) and **backup import**.
- **Types flow inward from schemas:** input types are `z.infer`red and mapped to canonical `domain/models`. Domain plausibility ranges from FITNESS_DOMAIN §8 are encoded as Zod refinements (e.g. `weight 0–500`, `reps 0–100`).
- **Validate once, at the edge.** Inner layers trust already-validated data; they do not re-validate defensively. Repositories assume valid domain models.
- **Import is the highest-risk edge:** every imported backup is fully Zod-validated against its declared schema version before a single row is written (§12).

---

## 11. Error Handling Strategy

- **Input/validation errors are values,** surfaced through RHF/Zod to the form UI — not exceptions.
- **Expected domain failures** (e.g. "no data in range") are represented as explicit states/`insufficient-data`, per FITNESS_DOMAIN — never thrown.
- **Infrastructure failures** (SQLite, FileSystem) throw **typed errors** (`DatabaseError`, `StorageError`) from `data/`/`core`, caught at the **feature boundary**, which renders a recoverable error state. The app never white-screens on a query failure.
- **Programmer errors** (invariant violations) throw and are allowed to surface in development; guarded by types so they're rare.

---

## 12. Local-First Integrity, Backup & Export

- **SQLite is the sole source of truth.** MMKV caches and in-memory stores are derived and disposable; on cold start the app rebuilds them from SQLite.
- **Migrations** are forward-only and versioned; the migration runner in `core/db` executes pending migrations inside the DB-ready gate before any feature renders. (Schema/migration detail owned by DATABASE.)
- **Export:** a `data/backup` service reads all tables via repositories, serializes to a **versioned document** (format owned by DATABASE), bundles progress-photo files, and hands off via the OS share sheet using `expo-file-system`.
- **Import:** the document is **Zod-validated against its version**, migrated if older, then written in a **single transaction** (all-or-nothing) so a failed import never corrupts existing history. Photos are restored to the document directory and re-linked by metadata.
- **Approved decision:** progress-photo **image bytes live in the filesystem** (app document directory); SQLite stores only metadata (id, date, angle, relative path). Backups bundle the files alongside the data document.

---

## 13. Performance Architecture

- **Compute off the render path:** analytics run in `useMemo`/derived selectors keyed on source-data versions, or are precomputed on write and cached in MMKV — never recomputed every frame.
- **Virtualized lists** (`FlatList`/FlashList) for workout history, meal logs, and measurement logs.
- **Index-backed queries** for date-ranged reads (indexes owned by DATABASE); range windows (FITNESS_DOMAIN §7) map to indexed `WHERE date >= ?` scans.
- **Transactional batch writes;** high-frequency editing (active session) stays in Zustand and hits SQLite only at checkpoints and on finish (§7.1).
- **Selective subscriptions:** hooks subscribe only to the tables they read, so a nutrition write doesn't re-render workout screens.

---

## 14. Module Boundaries & Enforcement

- **Feature isolation** and the **dependency rule** (§4) are enforced by an ESLint import-boundary rule (configuration owned by CODING_STANDARDS).
- **One repository per aggregate** (workouts, nutrition, body, photos, programs, foods, phases, settings/targets); repositories are the **only** modules containing SQL.
- **Promotion rule:** the moment two features need the same non-trivial logic or component, it moves down to `domain/` (logic) or `core/` (ui/infra). Features never reach sideways. Single-consumer logic stays in its feature (§9.1).

---

## 15. Testing Seams (identified here; strategy owned by DEVELOPMENT_WORKFLOW)

- **`domain/`** — pure functions; the primary, exhaustive unit-test target (every formula and edge case in FITNESS_DOMAIN §8).
- **`features/*/logic`** — feature-specific pure logic, unit-tested the same way.
- **`data/repositories`** — tested against a temporary/in-memory SQLite instance.
- **`features` hooks/components** — tested with React Native Testing Library where valuable.
- Because domain and feature logic are framework-free, the most valuable tests run without a native runtime.

---

## 16. AI Decision Rules (Architecture)

Binding rules for anyone extending the codebase:

1. **UI never touches SQL or repositories directly.** Components go through feature hooks/stores; only `data/repositories` contains SQL.
2. **Pure logic stays pure.** No React/RN/Expo/SQLite imports in `domain/` or `features/*/logic`. No side effects, no I/O. If you need data, receive it as an argument.
3. **SQLite is truth; MMKV is disposable.** Put anything backup-worthy or history-affecting in SQLite. Never store durable data in MMKV or Zustand.
4. **No cross-feature imports.** Shared code moves to `core/` or `domain/`. Features import down, never sideways.
5. **Keep routes thin.** `app/` files only wire navigation and render a feature screen; no business logic in route files.
6. **Validate at the edge with Zod, once.** All form and import input is validated at the boundary; inner layers trust validated data.
7. **Photos are files.** Image bytes live in the file system; SQLite stores only metadata and the relative path.
8. **Derived data is recomputed, not authoritative.** Never treat cached PRs/trends/insights as the source of truth; always recomputable from SQLite. Caches live in MMKV/memory and may be cleared safely.
9. **Round and format in the UI layer;** compute on raw values in `domain/`.
10. **Respect the responsibility matrix (§3).** Use each technology only for its lane; consult the owning document before bending a boundary.
11. **All writes that touch multiple rows run in a transaction;** imports are all-or-nothing.
12. **FITNESS_DOMAIN formulas have exactly one implementation, in `domain/`.** Never re-derive volume/e1RM/PR/adherence math in a feature, hook, or component.
13. **Don't over-abstract.** Single-consumer logic lives in its feature; promote on the second consumer, not before (§9.1). Prefer the simplest structure that keeps boundaries intact.
14. **The active workout must be crash-safe.** Any change to session handling preserves the draft-checkpoint + resume guarantee (§7.1).

---

## Changelog

- 2026-07-08 — v1 baseline frozen (five open decisions approved; F2 consistency amendment: Drizzle ORM + drizzle-kit and Victory Native XL added to the §3 matrix; Inter/Lucide noted as design assets).
- 2026-07-08 — Amendment (Phase 2, approved): §4 now explicitly permits composition-root wiring in `app/_layout` (mounting providers and injecting the `data/schema` migration bundle into the `core/db` gate). Documents existing architectural intent; no change to the dependency direction.
