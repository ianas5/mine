# CODING_STANDARDS.md

> **Status:** FROZEN (v1 baseline · 2026-07-08) · **Owner concern:** file-level code rules and conventions · **Depends on:** PRODUCT_PRINCIPLES (P22: readability over cleverness), ARCHITECTURE (macro structure; this document governs what a single file looks like inside it).
>
> These rules exist so that code written months apart reads as one hand's work. Where a rule is arbitrary (quote style), it is still binding — consistency is the value, not the particular choice.

---

## 1. TypeScript Rules

1. **Compiler:** `strict: true` plus `noUncheckedIndexedAccess`, `noImplicitOverride`, `noFallthroughCasesInSwitch`, `exactOptionalPropertyTypes`. Never weakened, globally or per-file.
2. **`any` is forbidden.** Use `unknown` + narrowing at true unknowns (Zod boundaries make these rare). `as` casts only at infrastructure edges (DB row mapping, JSON parse pre-validation) with a one-line justification comment. `@ts-ignore`/`@ts-expect-error` require a comment and are treated as open defects.
3. **Non-null assertions (`!`) forbidden** outside tests. Model absence honestly (`null` = not recorded, per FITNESS_DOMAIN §2.4) and handle it.
4. **Exported functions declare explicit return types.** Inference is fine for locals.
5. **Discriminated unions over boolean flags** for states (`MetricResult`, load types, insight tones are the house style). Exhaustive `switch` with `never` check on every union.
6. **`readonly` by default** for domain model fields and arrays crossing layer boundaries; mutation is an implementation detail inside a function, never an interface.
7. **Type naming:** no `I`/`T` prefixes. Domain models are bare nouns (`Workout`, `SetEntry`); component props `XxxProps`; DB rows `XxxRow`; Zod-inferred inputs `XxxInput`. Type aliases for semantic primitives: `IsoDate` (`YYYY-MM-DD` string), `EpochMs`, `Uuid` — never bare `string`/`number` for these in signatures.

## 2. Files & Naming

1. **One primary export per file**, file named after it: components/screens `PascalCase.tsx`; hooks `useXxx.ts`; everything else `camelCase.ts`. Expo Router files follow router conventions (lowercase route names) and contain nothing but the route wiring (ARCHITECTURE rule 5).
2. **Named exports only.** `export default` exists only where a framework demands it (Expo Router routes).
3. **Size discipline:** soft limit **200 lines** per file, **50 lines** per function. Passing it is a prompt to split by responsibility, not to reformat tighter (P22).
4. **Folder shape** is fixed by ARCHITECTURE §5. New file placement follows the state-ownership and layer tables — when unsure, place *lower* (toward `domain/`).
5. **Barrels are public APIs, not conveniences:** `core/ui`, `core/theme`, and each `domain/*` module export a small, curated surface from `index.ts`; everything unexported is private. Consumers depend on **behavior, not internal structure** — importing past a barrel (`@/domain/fitness/internal/…`) is a boundary violation (lint-enforced). A module whose "public API" exceeds ~10 exports is a prompt to split it. Repositories follow the same discipline: their object literal *is* the API; row types, mappers, and SQL stay private to `data/`. Features have **no barrels** — nothing imports a feature except `app/` routes (screens directly).
6. **Units live in names** exactly as in the schema: `weightKg`, `waistCm`, `waterMl`, `kcal`, `proteinG`, `durationMin`, `restSeconds`. A number without a unit in its name is a review flag.
7. **Booleans read as predicates** (`isWarmup`, `hasTarget`, `canDelete`); event props `onXxx`, handlers `handleXxx`; date fields `date: IsoDate`, timestamps `xxxAt: EpochMs`.

## 3. Imports & Boundaries

1. **Path aliases:** `@/core/*`, `@/domain/*`, `@/data/*`, `@/features/*`. No relative imports that climb more than one level (`../../` forbidden).
2. **Import order** (lint-enforced, auto-fixed): react/react-native → third-party → `@/core` → `@/domain` → `@/data` → `@/features` → relative. One blank line between groups.
3. **The ARCHITECTURE §4 dependency rules are lint rules,** not documentation: `eslint-plugin-boundaries` (or equivalent) encodes — `domain` imports nothing above it and no framework; `data` never imports `features`/`app`; features never import features; only `data/` imports `expo-sqlite`/Drizzle; only `core/storage` imports MMKV. A boundary violation is a build failure, not a warning.
4. **Token enforcement (DESIGN_SYSTEM §7):** `no-restricted-syntax` bans hex color literals and raw `fontSize`/`padding`/`margin` numbers outside `core/theme` + `core/ui`.

## 4. React & Component Patterns

1. **Function components only,** typed props (`function StatTile(props: StatTileProps)`), no `React.FC`.
2. **Component file order:** types → component → styles (`StyleSheet.create` at bottom) → local helpers. Hooks at top of the component; early-return the empty/insufficient/error/loading states *before* the main render (UI_UX rule 5 becomes code shape).
3. **No business logic in components.** A component may format, map, and dispatch — it may not compute domain values (ARCHITECTURE rule 12). If a component contains arithmetic beyond layout, it's misplaced.
4. **Custom hooks are the orchestration layer:** one hook per file, return a named object (`{ workouts, isLoading, refresh }`), never positional tuples. Data hooks subscribe to the change-bus via the shared subscription helper — no bespoke listeners in features.
5. **Memoization is a measured fix, not a habit** (ANALYTICS §7 posture): `useMemo`/`useCallback`/`memo` where a real re-render cost or dependency identity requires it, with the reason evident. FlashList/FlatList row components are the standing exception — always memoized.
6. **StyleSheet over inline styles;** inline only for genuinely dynamic values (progress widths). No style objects created per-render in lists.

## 5. State & Stores (Zustand)

1. One store per concern, in `features/*/stores/`, named `useXxxStore`. Store state is **ephemeral by definition** (ARCHITECTURE §6) — anything durable goes through a repository.
2. **Shape:** flat state + `actions` object created in the store; components select narrowly (`useSessionStore(s => s.restRemaining)`) — never subscribe to whole stores in render paths.
3. Stores may call repositories and domain functions inside actions; they contain **no formulas** (ARCHITECTURE rule 12) and no JSX.
4. The active-session store implements the checkpoint contract (ARCHITECTURE §7.1) via its actions — checkpointing is an action side-effect, never a component `useEffect`.

## 6. Validation & Schemas (Zod)

1. Form schemas live in `features/*/schemas/`, backup/import schemas in `data/backup/`. Named `xxxSchema`; inferred types `type XxxInput = z.infer<typeof xxxSchema>`.
2. **Plausibility ranges come from one constants module** — `domain/fitness/constants.ts` exports FITNESS_DOMAIN §8.13's ranges and §6.4/§6.5 thresholds by name (`WEIGHT_KG_MAX`, `TREND_MIN_POINTS`, `RECOMP_WEIGHT_STABLE_KG`…). Zod refinements, domain math, and tests all import these; **a threshold literal appearing inline anywhere is a defect** (single source of truth).
3. **Validate once, at the edge** (ARCHITECTURE §10); no defensive re-validation inside `domain/`/`data/`.

## 7. Domain Code

1. Pure functions, verb-named (`calculateSetVolume`, `classifyTrend`, `evaluateInsights`); data in, values out; "now" and settings are **parameters**, never imports (ANALYTICS rule 8).
2. Every function implementing a documented formula carries a JSDoc line citing its source: `/** Epley e1RM — FITNESS_DOMAIN §3.5 */`. The citation is the contract; a change without a doc amendment is a violation.
3. Calculators (ANALYTICS §3.4) are modules of functions, not classes. No classes anywhere except typed `Error` subclasses (`DatabaseError`, `StorageError` — ARCHITECTURE §11).
4. **Comments explain *why*, never *what*.** Dead code is deleted, not commented out; `TODO(scope): reason` is the only sanctioned marker and is expected to be short-lived.

## 8. Data Layer

1. Repositories: object-literal modules (`export const workoutRepository = {...}`), returning **domain models via mappers** — a `XxxRow` type never crosses out of `data/`.
2. Drizzle query builder over raw SQL strings; raw SQL only where the builder genuinely cannot express the query, isolated and commented.
3. Every multi-row write wrapped in a transaction; every write emits its change-bus event as the last step — both are review checklist items, not conventions to remember.

## 9. Formatting & Tooling

1. **Prettier owns formatting** — nobody argues with it: 2-space indent, single quotes, semicolons, trailing commas `all`, `printWidth: 100`. Runs on save and in CI.
2. **ESLint:** `typescript-eslint` strict + `react-hooks` + boundaries + the restrictions above. Zero-warning policy: warnings fail CI (a tolerated warning is a future bug with seniority).
3. Both configured once at repo root; per-file disables require a reason comment and are review flags.

## 10. Testing Conventions

*(What to test and when is DEVELOPMENT_WORKFLOW's; how tests read is standardized here.)*

1. Colocated `xxx.test.ts` beside the unit; one runner for everything: **Jest via `jest-expo`**.
2. Structure: `describe` = unit under test, `it` = one behavior in plain English (`it('excludes warm-up sets from volume')`); arrange–act–assert with a blank line between phases; no logic in tests (no loops/conditionals around expectations — table-driven via `it.each`).
3. Domain tests use the FITNESS_DOMAIN edge-case list (§8) as their checklist; fixtures are tiny builders (`makeSet({ warmup: true })`), never 200-line JSON blobs.

## 11. AI Decision Rules (Coding Standards)

1. **Readable beats clever, boring beats novel** (P22). If a construct needs a comment to explain *how it works* (not why), rewrite it.
2. **Follow the file recipes:** naming, order, size limits, and import rules above are mechanical — apply them without exception; propose amendments here rather than deviating locally.
3. **No inline domain literals.** Any formula constant or plausibility bound comes from `domain/fitness/constants.ts` with its FITNESS_DOMAIN citation.
4. **Never weaken tooling to make code pass:** no strictness downgrades, no lint-rule disables, no `any`, no `!` — fix the code, or amend the standard explicitly.
5. **Types tell the truth:** if a value can be absent, its type says so; if a state is a union, model the union — never encode states in comments or conventions.
6. **When two styles are both fine, pick the one already in the codebase.** Consistency outranks preference; this document outranks both.
7. **New patterns enter this document before the codebase.** A second way of writing stores, hooks, or repositories is a defect even if it's better — amend first, then migrate uniformly.
8. **Documentation leads implementation — always.** If implementing a feature requires changing an accepted project document, **stop implementation first.** Update the document, obtain approval, and only then continue. Code that lands ahead of its documentation is wrong even when it works (P22; PRODUCT_PRINCIPLES amendment rule).
9. **The Rule of Two.** The first duplication is fine; the **second occurrence of the same logic triggers refactoring now** — extract it (per the ARCHITECTURE §9.1 promotion rule) before it spreads. The inverse binds equally: **never create an abstraction before the second real use case exists.** One consumer = no abstraction; two consumers = mandatory one.
10. **Prefer the platform before adding dependencies.** If React Native, Expo, TypeScript, Drizzle, Zustand, Zod, RHF, or the existing stack solves the problem well, adding a library for it is forbidden. Every new dependency passes the PRODUCT_PRINCIPLES dependency test (offline-forever, P19) **plus**: actively maintained, understood well enough to debug, and cheaper than writing the 50 lines ourselves. The approved stack list lives in ARCHITECTURE §3 — a new dependency amends it first (rule 8).
11. **Leave it better than you found it.** Every pull request leaves the codebase slightly better — a clearer name, a dead file removed, a missing test added — within the files it already touches. Small continuous improvement beats occasional massive refactors; "big cleanup later" is how personal projects die (P4, P22). Scope discipline still applies: improve what you touch, don't turn a fix into a rewrite.

---

## Changelog

- 2026-07-08 — v1 baseline frozen (five approved refinements applied: documentation-first rule, Rule of Two, platform-first dependencies, public-API barrels, leave-it-better).
